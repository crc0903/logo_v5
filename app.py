
import streamlit as st
from PIL import Image
import os
import io
import math
from pptx import Presentation
from pptx.util import Inches

PRELOADED_LOGO_DIR = "preloaded_logos"

def load_preloaded_logos():
    logos = {}
    if not os.path.exists(PRELOADED_LOGO_DIR):
        os.makedirs(PRELOADED_LOGO_DIR)
    for file in os.listdir(PRELOADED_LOGO_DIR):
        if file.lower().endswith((".png", ".jpg", ".jpeg", ".webp")):
            name = os.path.splitext(file)[0]
            image = Image.open(os.path.join(PRELOADED_LOGO_DIR, file)).convert("RGBA")
            logos[name] = image
    return logos

def resize_to_fit(image, target_width, target_height):
    img_w, img_h = image.size
    ratio = min(target_width / img_w, target_height / img_h)
    new_size = (int(img_w * ratio), int(img_h * ratio))
    return image.resize(new_size, Image.LANCZOS)

def create_logo_slide(prs, logos, canvas_width_in, canvas_height_in, logos_per_row):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    canvas_width_px = int(canvas_width_in * 96)
    canvas_height_px = int(canvas_height_in * 96)

    logo_count = len(logos)
    cols = logos_per_row if logos_per_row else max(1, round(math.sqrt(logo_count * canvas_width_in / canvas_height_in)))
    rows = math.ceil(logo_count / cols)

    cell_width = canvas_width_px / cols
    cell_height = canvas_height_px / rows

    left_margin = Inches((10 - canvas_width_in) / 2)
    top_margin = Inches((7.5 - canvas_height_in) / 2)

    for idx, logo in enumerate(logos):
        col = idx % cols
        row = idx // cols
        resized = resize_to_fit(logo, cell_width, cell_height)

        img_stream = io.BytesIO()
        resized.save(img_stream, format="PNG")
        img_stream.seek(0)

        left = left_margin + Inches(col * (canvas_width_in / cols))
        top = top_margin + Inches(row * (canvas_height_in / rows))
        slide.shapes.add_picture(img_stream, left, top, width=Inches(resized.width / 96), height=Inches(resized.height / 96))

st.title("Logo Grid PowerPoint Exporter")
st.markdown("Upload logos or use preloaded ones below:")

uploaded_files = st.file_uploader("Upload logos", type=["png", "jpg", "jpeg", "webp"], accept_multiple_files=True)
preloaded = load_preloaded_logos()
selected_preloaded = st.multiselect("Select preloaded logos", options=list(preloaded.keys()))

canvas_width_in = st.number_input("Grid width (inches)", min_value=1.0, max_value=20.0, value=10.0)
canvas_height_in = st.number_input("Grid height (inches)", min_value=1.0, max_value=20.0, value=7.5)
logos_per_row = st.number_input("Logos per row (optional)", min_value=0, max_value=50, value=0)

if st.button("Generate PowerPoint"):
    images = []

    if uploaded_files:
        for f in uploaded_files:
            image = Image.open(f).convert("RGBA")
            images.append(image)

    for name in selected_preloaded:
        images.append(preloaded[name])

    if not images:
        st.warning("Please upload or select logos.")
    else:
        prs = Presentation()
        create_logo_slide(prs, images, canvas_width_in, canvas_height_in, logos_per_row if logos_per_row > 0 else None)

        output = io.BytesIO()
        prs.save(output)
        output.seek(0)

        st.success("PowerPoint created!")
        st.download_button("Download .pptx", output, file_name="logo_grid.pptx")
