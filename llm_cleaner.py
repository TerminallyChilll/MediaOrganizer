"""
LLM-powered title cleaning for the Media Organizer.
Supports: Gemini, OpenAI, and Ollama (local).
Uses only stdlib (urllib + json) — no pip installs required.
"""

import json, os, urllib.request, urllib.error, ssl

LLM_CONFIG_FILE = ".media_llm_config.json"

# ─── Config Caching ──────────────────────────────────────────────────

def load_llm_config():
    if os.path.exists(LLM_CONFIG_FILE):
        try:
            with open(LLM_CONFIG_FILE, 'r') as f:
                return json.load(f)
        except Exception: pass
    return {}

def save_llm_config(config):
    try:
        with open(LLM_CONFIG_FILE, 'w') as f:
            json.dump(config, f, indent=2)
    except Exception: pass


# ─── Shared Prompt ────────────────────────────────────────────────────

def build_prompt(filenames):
    """Build a prompt that asks the LLM to clean media filenames."""
    names_list = "\n".join(f"{i+1}. {name}" for i, name in enumerate(filenames))
    return f"""Clean these media filenames. For each one, extract the movie/show title, year, and quality.

IMPORTANT: Return ONLY a JSON array. Do NOT write code. Do NOT explain anything.

Example input: "The.Matrix.1999.1080p.BluRay.x264-GROUP"
Example output: [{{"original": "The.Matrix.1999.1080p.BluRay.x264-GROUP", "title": "The Matrix", "year": "1999", "quality": "1080p"}}]

Rules:
- "original" must be the EXACT input filename unchanged
- "title" is the clean movie/show name in Title Case, no dots/underscores, no codec/source/group junk
- "year" is the 4-digit year if found, otherwise ""
- "quality" is the resolution (1080p, 720p, 4K, 2160p) if found, otherwise ""
- Do NOT include file extensions (.mkv, .mp4) in the title

Filenames to clean:
{names_list}

Return ONLY the JSON array, nothing else:"""


# ─── API Callers ──────────────────────────────────────────────────────

def _make_request(url, data, headers, timeout=60):
    """Make an HTTP POST request and return the response body as string."""
    body = json.dumps(data).encode('utf-8')
    req = urllib.request.Request(url, data=body, headers=headers, method='POST')
    
    # Create SSL context that works on Windows
    ctx = ssl.create_default_context()
    
    try:
        with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
            return json.loads(resp.read().decode('utf-8'))
    except urllib.error.HTTPError as e:
        error_body = e.read().decode('utf-8') if e.fp else ''
        raise Exception(f"HTTP {e.code}: {error_body[:200]}")  # type: ignore
    except urllib.error.URLError as e:
        raise Exception(f"Connection failed: {e.reason}")


def call_gemini(filenames, api_key, model="gemini-2.0-flash"):
    """Call Google Gemini API."""
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
    data = {
        "contents": [{"parts": [{"text": build_prompt(filenames)}]}],
        "generationConfig": {"temperature": 0.1, "maxOutputTokens": 8192}
    }
    headers = {"Content-Type": "application/json"}
    
    resp = _make_request(url, data, headers, timeout=120)
    text = resp['candidates'][0]['content']['parts'][0]['text']
    return _parse_llm_response(text, filenames)


def call_openai(filenames, api_key, model="gpt-4o-mini"):
    """Call OpenAI API."""
    url = "https://api.openai.com/v1/chat/completions"
    data = {
        "model": model,
        "messages": [
            {"role": "system", "content": "You are a media filename parser. Return only valid JSON."},
            {"role": "user", "content": build_prompt(filenames)}
        ],
        "temperature": 0.1,
        "max_tokens": 8192
    }
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    
    resp = _make_request(url, data, headers, timeout=120)
    text = resp['choices'][0]['message']['content']
    return _parse_llm_response(text, filenames)


def list_ollama_models(base_url="http://localhost:11434"):
    """Query local or remote Ollama for installed models. Returns list of model names."""
    try:
        url = f"{base_url}/api/tags"
        req = urllib.request.Request(url, method='GET')
        with urllib.request.urlopen(req, timeout=5) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            models = [m['name'] for m in data.get('models', [])]
            return models
    except Exception:
        return []

def call_ollama(filenames, model="llama3", base_url="http://localhost:11434"):
    """Call local or remote Ollama API using chat endpoint for better results."""
    url = f"{base_url}/api/chat"
    data = {
        "model": model,
        "messages": [
            {"role": "system", "content": "You are a JSON API. You receive media filenames and return a JSON array with clean titles. Never write code. Never explain. Return ONLY valid JSON."},
            {"role": "user", "content": build_prompt(filenames)}
        ],
        "stream": False,
        "options": {"temperature": 0.1}
    }
    headers = {"Content-Type": "application/json"}
    
    resp = _make_request(url, data, headers, timeout=300)
    text = resp.get('message', {}).get('content', '')
    return _parse_llm_response(text, filenames)


# ─── Response Parsing ─────────────────────────────────────────────────

def _parse_llm_response(text, original_filenames):
    """Parse the LLM's JSON response into a dict of {original: {title, year, quality}}."""
    if not text or not text.strip():
        print("      ⚠️ LLM returned empty response")
        return {}
    
    raw_text = text.strip()
    
    # Strip markdown code fences if present
    if raw_text.startswith('```'):
        lines = raw_text.split('\n')
        raw_text = '\n'.join(lines[1:])
        if raw_text.rstrip().endswith('```'):
            raw_text = raw_text.rstrip()[:-3]  # type: ignore
        raw_text = raw_text.strip()
    
    # Try direct JSON parse first
    results = None
    try:
        results = json.loads(raw_text)
    except json.JSONDecodeError:
        pass
    
    # Try to extract JSON array from text
    if results is None:
        start = raw_text.find('[')
        end = raw_text.rfind(']')
        if start != -1 and end != -1 and end > start:
            json_candidate = raw_text[start:end+1]  # type: ignore
            try:
                results = json.loads(json_candidate)
            except json.JSONDecodeError:
                # Try fixing common local model issues: trailing commas, single quotes
                fixed = json_candidate.replace("'", '"')
                # Remove trailing commas before ] or }
                import re as _re
                fixed = _re.sub(r',\s*]', ']', fixed)
                fixed = _re.sub(r',\s*}', '}', fixed)
                try:
                    results = json.loads(fixed)
                except Exception:
                    pass
    
    if results is None:
        # Show first 300 chars of the response for debugging
        preview = raw_text[:300].replace('\n', ' ')  # type: ignore
        print(f"      ⚠️ Could not parse LLM response as JSON. Preview: {preview}...")
        return {}
    
    if not isinstance(results, list):
        print(f"      ⚠️ LLM returned {type(results).__name__} instead of a list")
        return {}
    
    # Build lookup dict - be flexible with field names
    cleaned = {}
    for item in results:
        if not isinstance(item, dict):
            continue
        # Try multiple possible field names the LLM might use
        original = item.get('original', '') or item.get('filename', '') or item.get('input', '') or item.get('name', '')
        title = item.get('title', '') or item.get('clean_title', '') or item.get('cleaned', '')
        year = item.get('year', '') or item.get('release_year', '')
        quality = item.get('quality', '') or item.get('resolution', '')
        
        if original and title:
            cleaned[original] = {
                'title': str(title).strip(),
                'year': str(year).strip() if year else '',
                'quality': str(quality).strip() if quality else ''
            }
    
    if not cleaned and results:
        # We parsed JSON but couldn't match fields - show what we got
        sample = results[0] if results else {}
        print(f"      ⚠️ Parsed {len(results)} items but couldn't extract data. Sample keys: {list(sample.keys()) if isinstance(sample, dict) else 'N/A'}")
    
    return cleaned


# ─── Main Entry Point ─────────────────────────────────────────────────

def clean_titles_with_llm(filenames, provider, api_key=None, model=None, ollama_url=None):
    """
    Clean a list of filenames using the specified LLM provider.
    Returns dict: {original_filename: {title, year, quality}}
    """
    if not filenames:
        return {}
    
    # Process in batches — smaller for local models
    batch_size = 10 if provider == 'ollama' else 25
    all_results = {}
    
    for i in range(0, len(filenames), batch_size):
        batch = filenames[i:i+batch_size]  # type: ignore
        batch_num = (int(i) // int(batch_size)) + 1
        total_batches = (len(filenames) + int(batch_size) - 1) // int(batch_size)
        
        if total_batches > 1:
            print(f"   🤖 Processing batch {batch_num}/{total_batches} ({len(batch)} names)...")
        
        try:
            if provider == 'gemini':
                results = call_gemini(batch, api_key, model or "gemini-2.0-flash")
            elif provider == 'openai':
                results = call_openai(batch, api_key, model or "gpt-4o-mini")
            elif provider == 'ollama':
                results = call_ollama(batch, model or "llama3", base_url=ollama_url or "http://localhost:11434")
            else:
                raise Exception(f"Unknown provider: {provider}")
            
            all_results.update(results)
        except Exception as e:
            print(f"   ❌ LLM error on batch {batch_num}: {e}")
    
    return all_results
