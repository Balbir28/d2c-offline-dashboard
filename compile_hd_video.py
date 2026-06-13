import os
import glob
import cv2
import shutil

def compile_hd_video():
    brain_dir = "/Users/balbaasaur/.gemini/antigravity-ide/brain/d69807e3-4a91-495c-9f6f-3c324bb4b0f8"
    frames_dir = os.path.join(brain_dir, "scratch", "hd_frames")
    
    # Grab all generated frame images in order
    pattern = os.path.join(frames_dir, "frame_*.png")
    frame_files = sorted(glob.glob(pattern))
    
    if not frame_files:
        print("Error: No frame files found to compile!")
        return
        
    print(f"Compiling {len(frame_files)} frames into Full HD video...")
    
    # Video details
    width = 1920
    height = 1080
    fps = 10
    
    output_path = os.path.join(brain_dir, "v1.mp4")
    fourcc = cv2.VideoWriter_fourcc(*'mp4v') # MPEG-4 standard codec
    
    video = cv2.VideoWriter(output_path, fourcc, fps, (width, height))
    
    if not video.isOpened():
        print("Error: Could not open VideoWriter.")
        return
        
    for idx, path in enumerate(frame_files):
        if idx % 100 == 0:
            print(f"Stitching frame {idx}/{len(frame_files)}...")
            
        img = cv2.imread(path)
        if img is None:
            print(f"Warning: Could not read frame {path}. Skipping...")
            continue
            
        # Resize to exactly 1920x1080 if not already (safety check)
        h, w, c = img.shape
        if w != width or h != height:
            img = cv2.resize(img, (width, height), interpolation=cv2.INTER_CUBIC)
            
        video.write(img)
        
    video.release()
    print(f"Successfully compiled HD MP4 video at: {output_path}")
    
    # Copy to user's downloads folder as 'v1.mp4'
    downloads_dir = "/Users/balbaasaur/Downloads"
    if os.path.exists(downloads_dir):
        dest_path = os.path.join(downloads_dir, "v1.mp4")
        shutil.copy(output_path, dest_path)
        print(f"Successfully copied video to Downloads at: {dest_path}")
        
    # Clean up frames directory to save space
    try:
        shutil.rmtree(frames_dir)
        print("Cleaned up temporary frames folder.")
    except Exception as e:
        print(f"Failed to delete temporary frames folder: {e}")

if __name__ == "__main__":
    compile_hd_video()
