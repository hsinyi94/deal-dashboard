"""Extract ALL images from page 1 of PDFs likely to contain Amazon wordmark."""
import os
import fitz  # PyMuPDF

CANDIDATES = [
    "PointsDEAL_BulkUpload_2025Nov_CN.pdf",
    "TWGS Branding Visual Identity Guideline_251124.pdf",
    "2026日亚企业购品牌专场活动介绍-普通展位.pdf",
    "2026日亚企业购品牌专场活动介绍-普通展位_繁體.pdf",
    "MDE4/2026_MDE4_3P_CN.pdf",
]

outdir = "extracted_logos"
os.makedirs(outdir, exist_ok=True)

for pdf in CANDIDATES:
    if not os.path.exists(pdf):
        continue
    doc = fitz.open(pdf)
    base = os.path.splitext(os.path.basename(pdf))[0]
    print(f"\n=== {pdf} : {len(doc)} pages ===")
    for pno in range(min(2, len(doc))):
        page = doc[pno]
        imgs = page.get_images(full=True)
        print(f"  page {pno+1}: {len(imgs)} images")
        for img_idx, img in enumerate(imgs):
            xref = img[0]
            base_img = doc.extract_image(xref)
            ext = base_img["ext"]
            w, h = base_img["width"], base_img["height"]
            out = os.path.join(outdir, f"{base}_p{pno+1}_{img_idx}.{ext}")
            with open(out, "wb") as f:
                f.write(base_img["image"])
            print(f"    [{img_idx}] {w}x{h} -> {out}")
    doc.close()
