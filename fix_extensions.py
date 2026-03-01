"""
Emergency Recovery Script: Restore stripped file extensions.
Scans media folders for files missing video extensions and restores them
by reading the file's magic bytes to detect the actual format.
"""
import os, sys, shutil

MAGIC_BYTES = {
    b'\x1a\x45\xdf\xa3': '.mkv',       # Matroska / MKV / WebM
    b'\x00\x00\x00': '.mp4',            # MP4/M4V (ftyp box starts at byte 4)
    b'RIFF': '.avi',                     # AVI
    b'\x00\x00\x01\xba': '.mpg',        # MPEG
    b'\x00\x00\x01\xb3': '.mpg',        # MPEG
    b'\x30\x26\xb2\x75': '.wmv',        # WMV/ASF
    b'FLV\x01': '.flv',                 # FLV
}

VIDEO_EXTENSIONS = {'.mkv', '.mp4', '.avi', '.mov', '.m4v', '.wmv', '.flv', '.webm', '.ts', '.mpg', '.mpeg', '.m2ts'}

def detect_extension(filepath):
    """Read file header bytes to determine the video format."""
    try:
        with open(filepath, 'rb') as f:
            header = f.read(12)
        
        # Check MKV/WebM (Matroska)
        if header.startswith(b'\x1a\x45\xdf\xa3'):
            return '.mkv'
        
        # Check MP4/M4V (look for 'ftyp' at byte 4)
        if len(header) >= 8 and header.startswith(b'ftyp', 4):
            return '.mp4'
        
        # Check AVI
        if header.startswith(b'RIFF') and len(header) >= 12 and header.startswith(b'AVI ', 8):
            return '.avi'
        
        # Check MPEG
        if header.startswith(b'\x00\x00\x01'):
            return '.mpg'
        
        # Check WMV
        if header.startswith(b'\x30\x26\xb2\x75'):
            return '.wmv'
        
        # Check FLV
        if header.startswith(b'FLV\x01'):
            return '.flv'
        
        # Return None for unrecognized formats instead of forcing .mkv
        return None
    except Exception:
        return None

def scan_and_fix(folder_path, dry_run=True):
    """Walk through folders finding files without video extensions."""
    fixes = []
    
    for dirpath, dirnames, filenames in os.walk(folder_path):
        for filename in filenames:
            filepath = os.path.join(dirpath, filename)
            _, ext = os.path.splitext(filename)
            
            # Skip files that already have a video extension
            if ext.lower() in VIDEO_EXTENSIONS:
                continue
            
            # Skip non-video files (like .nfo, .srt, .jpg, etc.)
            known_non_video = {'.nfo', '.srt', '.sub', '.idx', '.jpg', '.jpeg', '.png', '.txt', '.nzb', '.bat', '.sh', '.py', '.json', '.xlsx', '.xls'}
            if ext.lower() in known_non_video:
                continue
            
            # This file has no extension or an unrecognized one - check if it's a video
            detected_ext = detect_extension(filepath)
            if detected_ext:
                new_name = filename + detected_ext
                new_path = os.path.join(dirpath, new_name)
                fixes.append((filepath, new_path, detected_ext))
    
    return fixes

def main():
    print("=" * 70)
    print("🔧 FILE EXTENSION RECOVERY TOOL")
    print("=" * 70)
    
    path = input("\nEnter the media folder path to scan: ").strip()
    if path.startswith('"') and path.endswith('"'):
        path = path.strip('"')
    
    if not os.path.exists(path):
        print(f"❌ Path does not exist: {path}")
        return
    
    print(f"\n🔍 Scanning: {path}")
    fixes = scan_and_fix(path)
    
    if not fixes:
        print("\n✅ No files found that need extension recovery!")
        return
    
    print(f"\n📋 Found {len(fixes)} file(s) that need extensions restored:\n")
    for i, (old_path, new_path, ext) in enumerate(fixes):
        if i >= 20: break
        old_name = os.path.basename(old_path)
        new_name = os.path.basename(new_path)
        print(f"  {old_name}  →  {new_name}")
    
    if len(fixes) > 20:
        print(f"  ... and {len(fixes) - 20} more.")
    
    confirm = input(f"\nProceed with restoring {len(fixes)} extension(s)? (y/n): ").strip().upper()
    if confirm not in ['Y', 'YES']:
        print("Cancelled.")
        return
    
    success = 0
    for old_path, new_path, ext in fixes:
        try:
            os.rename(old_path, new_path)
            success += 1
        except Exception as e:
            print(f"  ❌ Failed: {os.path.basename(old_path)}: {e}")
    
    print(f"\n✅ Restored extensions on {success}/{len(fixes)} files.")

if __name__ == '__main__':
    main()
