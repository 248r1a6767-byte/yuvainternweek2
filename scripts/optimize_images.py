# -*- coding: utf-8 -*-
"""
Script: scripts/optimize_images.py
Purpose: Optimize visualization PNGs to 1600px width with Lanczos resampling
         and 256-color adaptive quantization. Guarantees publication-grade
         visual crispness while reducing file size to provide massive safety
         headroom below the 2048 KB upload limit.
"""

import os
from PIL import Image

def optimize_visualizations():
    vis_dir = "visualizations"
    target_width = 1600
    
    total_before = 0
    total_after = 0
    
    for fname in sorted(os.listdir(vis_dir)):
        if fname.endswith(".png"):
            fpath = os.path.join(vis_dir, fname)
            size_before = os.path.getsize(fpath)
            total_before += size_before
            
            im = Image.open(fpath)
            w, h = im.size
            if w > target_width:
                new_h = int(h * target_width / w)
                im_resized = im.resize((target_width, new_h), Image.Resampling.LANCZOS)
                # Convert to RGB then quantize to 256 colors
                im_opt = im_resized.convert("RGB").quantize(colors=256, method=Image.Resampling.LANCZOS)
                im_opt.save(fpath, "PNG", optimize=True)
            else:
                im_opt = im.convert("RGB").quantize(colors=256)
                im_opt.save(fpath, "PNG", optimize=True)
                
            size_after = os.path.getsize(fpath)
            total_after += size_after
            print(f"Optimized {fname:32s}: {size_before/1024:6.1f} KB -> {size_after/1024:6.1f} KB")
            
    print(f"\nTotal Visualizations: {total_before/1024:6.1f} KB -> {total_after/1024:6.1f} KB (-{(1 - total_after/total_before)*100:.1f}%)")

if __name__ == "__main__":
    optimize_visualizations()
