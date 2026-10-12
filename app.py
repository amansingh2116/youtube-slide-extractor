import streamlit as st
import cv2
from PIL import Image
import imagehash
import pytesseract
import os
import re
import shutil
import tempfile
import math
import json
import numpy as np
import base64
import zipfile
import io
import urllib.parse
import urllib.request
import urllib.error
import yt_dlp as youtube_dl
import streamlit.components.v1 as components
from pptx import Presentation
from pptx.util import Inches
from pptx.util import Pt
from pptx.dml.color import RGBColor

TESSERACT_WINDOWS_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


if os.path.exists(TESSERACT_WINDOWS_PATH):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_WINDOWS_PATH


def is_tesseract_available():
    return shutil.which("tesseract") or os.path.exists(TESSERACT_WINDOWS_PATH)


DEFAULT_OCR_SETTINGS = {
    "method": "local",
    "api_key": "",
    "engine": "3",
    "language": "auto",
    "is_table": False,
    "crop_mode": "auto",
}


# Function to download the video from YouTube
def download_video(url, output_path='video.mp4'):
    video_id_match = re.search(r'(?:v=|youtu\.be/|shorts/)([A-Za-z0-9_-]{11})', url)
    if not video_id_match:
        st.error("Please enter a complete YouTube URL with an 11-character video ID.")
        return None

    try:
        ydl_opts = {
            'format': 'bv*+ba/b',
            'outtmpl': output_path,
            'merge_output_format': 'mp4',
            'noplaylist': True,
        }
        with youtube_dl.YoutubeDL(ydl_opts) as ydl:
            ydl.download([url])
        return output_path
    except Exception as e:
        st.error(f"Video download failed: {e}")
        print(f"Error: {e}")
        return None


def save_uploaded_video(uploaded_file, output_path):
    with open(output_path, "wb") as video_file:
        video_file.write(uploaded_file.getbuffer())
    return output_path


def get_video_metadata(cap):
    fps = cap.get(cv2.CAP_PROP_FPS) or 0
    frame_count = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0)
    duration_seconds = frame_count / fps if fps > 0 and frame_count > 0 else 0
    return fps, frame_count, duration_seconds


def get_video_metadata_from_path(video_path):
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        return 0, 0, 0

    metadata = get_video_metadata(cap)
    cap.release()
    return metadata


def get_video_frame_at_time(video_path, time_seconds):
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        return None

    fps, frame_count, _ = get_video_metadata(cap)
    cap.set(cv2.CAP_PROP_POS_MSEC, max(0, time_seconds) * 1000)
    ret, frame = cap.read()

    if not ret and fps > 0 and frame_count > 0:
        target_frame = min(int(time_seconds * fps), frame_count - 1)
        cap.set(cv2.CAP_PROP_POS_FRAMES, target_frame)
        ret, frame = cap.read()

    cap.release()
    return frame if ret else None


def apply_frame_crop(frame, crop_settings=None):
    if not crop_settings or not crop_settings.get("enabled"):
        return frame

    height, width = frame.shape[:2]
    left = int(width * crop_settings.get("left", 0) / 100)
    right = int(width * crop_settings.get("right", 100) / 100)
    top = int(height * crop_settings.get("top", 0) / 100)
    bottom = int(height * crop_settings.get("bottom", 100) / 100)

    left = max(0, min(left, width - 1))
    right = max(left + 1, min(right, width))
    top = max(0, min(top, height - 1))
    bottom = max(top + 1, min(bottom, height))

    return frame[top:bottom, left:right]


# Function to extract frames from the video at specific intervals
def extract_frames(
    video_path,
    sample_interval_seconds=5,
    max_samples=400,
    start_time_seconds=0,
    end_time_seconds=None,
    frame_crop_settings=None,
):
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        st.error("Could not open the video file. Please try another video.")
        return []

    fps, frame_count, duration_seconds = get_video_metadata(cap)
    frames = []
    start_time_seconds = max(0, start_time_seconds or 0)
    if end_time_seconds and end_time_seconds > start_time_seconds:
        end_time_seconds = min(end_time_seconds, duration_seconds) if duration_seconds else end_time_seconds
    else:
        end_time_seconds = duration_seconds if duration_seconds else None

    if fps <= 0 or frame_count <= 0:
        frame_interval = 150
        frame_id = 0

        while cap.isOpened() and len(frames) < max_samples:
            ret, frame = cap.read()
            if not ret:
                break
            if frame_id % frame_interval == 0:
                frames.append(apply_frame_crop(frame, frame_crop_settings))
            frame_id += 1

        cap.release()
        return frames

    selected_duration = max(0, (end_time_seconds or duration_seconds) - start_time_seconds)
    sample_count = min(max_samples, math.ceil(selected_duration / sample_interval_seconds))
    if sample_count <= 0:
        cap.release()
        return []

    st.caption(
        f"Video length: {duration_seconds / 60:.1f} minutes. "
        f"Selected range: {start_time_seconds / 60:.1f}-{(end_time_seconds or duration_seconds) / 60:.1f} minutes. "
        f"Checking {sample_count} sampled frame(s)."
    )
    progress = st.progress(0)

    for sample_index in range(sample_count):
        target_seconds = start_time_seconds + (sample_index * sample_interval_seconds)
        if end_time_seconds and target_seconds > end_time_seconds:
            break

        target_msec = target_seconds * 1000
        target_frame = min(int(target_seconds * fps), frame_count - 1)

        cap.set(cv2.CAP_PROP_POS_MSEC, target_msec)
        ret, frame = cap.read()

        if not ret:
            cap.set(cv2.CAP_PROP_POS_FRAMES, target_frame)
            ret, frame = cap.read()

        if ret:
            frames.append(apply_frame_crop(frame, frame_crop_settings))

        progress.progress((sample_index + 1) / sample_count)

    cap.release()
    return frames

# Function to get unique slides based on perceptual hashing
def get_unique_slides(frames, hash_difference_threshold=5):
    hashes = []
    unique_frames = []
    progress = st.progress(0)
    
    for idx, frame in enumerate(frames):
        pil_image = Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
        pil_image.thumbnail((640, 360))
        hash_val = imagehash.phash(pil_image)
        
        # Only add frames that are sufficiently different from previous frames
        if all(abs(hash_val - h) > hash_difference_threshold for h in hashes):
            unique_frames.append(frame)
            hashes.append(hash_val)

        progress.progress((idx + 1) / len(frames))
    
    return unique_frames

def crop_to_slide_region(pil_image):
    image = np.array(pil_image.convert("RGB"))
    gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
    mask = cv2.inRange(gray, 185, 255)
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (21, 21))
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel)

    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    image_area = image.shape[0] * image.shape[1]
    best_box = None
    best_area = 0

    for contour in contours:
        x, y, width, height = cv2.boundingRect(contour)
        area = width * height
        if area > best_area and area > image_area * 0.12 and width > image.shape[1] * 0.25:
            best_box = (x, y, width, height)
            best_area = area

    if best_box is None:
        return pil_image.convert("RGB")

    x, y, width, height = best_box
    region_aspect = width / max(height, 1)
    original_aspect = image.shape[1] / max(image.shape[0], 1)

    # Screen recordings often include a tall PDF/viewer canvas around a 16:9 slide.
    # If the detected bright region is too tall, crop inward to the presentation area.
    if original_aspect > 1.6 and region_aspect < 1.45:
        target_aspect = 16 / 9
        target_width = width
        target_height = int(target_width / target_aspect)

        if target_height > height:
            target_height = height
            target_width = int(target_height * target_aspect)

        x = x + max(0, (width - target_width) // 2)
        y = y + max(0, int((height - target_height) * 0.35))
        width = target_width
        height = target_height

    padding = 8
    left = max(0, x - padding)
    top = max(0, y - padding)
    right = min(image.shape[1], x + width + padding)
    bottom = min(image.shape[0], y + height + padding)
    return pil_image.crop((left, top, right, bottom)).convert("RGB")


def get_ocr_source_image(pil_image, ocr_settings=None):
    settings = {**DEFAULT_OCR_SETTINGS, **(ocr_settings or {})}
    if settings.get("crop_mode") == "manual":
        return pil_image.convert("RGB")

    return crop_to_slide_region(pil_image)


def prepare_image_for_ocr(pil_image, ocr_settings=None):
    source_image = get_ocr_source_image(pil_image, ocr_settings)
    image = np.array(source_image)
    gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)

    scale = max(2.0, min(4.0, 1800 / max(gray.shape[1], 1)))
    gray = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    gray = cv2.fastNlMeansDenoising(gray, h=10)
    gray = cv2.equalizeHist(gray)

    return Image.fromarray(gray)


def score_ocr_text(text):
    if not text:
        return 0

    readable_chars = sum(char.isalnum() for char in text)
    words = re.findall(r"[A-Za-z0-9][A-Za-z0-9&().,:%/-]*", text)
    long_words = sum(1 for word in words if len(word) >= 3)
    noise_penalty = text.count("|") + text.count("_") + text.count("~")
    return readable_chars + (long_words * 8) - noise_penalty


def run_local_tesseract_ocr(pil_image, ocr_settings=None):
    if not is_tesseract_available():
        return None

    source_image = get_ocr_source_image(pil_image, ocr_settings)
    prepared_image = prepare_image_for_ocr(pil_image, ocr_settings)
    local_variants = [
        prepared_image,
        source_image.convert("L").resize(
            (source_image.width * 2, source_image.height * 2),
            Image.Resampling.LANCZOS,
        ),
    ]
    configs = [
        "--oem 3 --psm 6 -c preserve_interword_spaces=1",
        "--oem 3 --psm 4 -c preserve_interword_spaces=1",
        "--oem 3 --psm 11",
    ]
    results = []

    for variant in local_variants:
        for config in configs:
            try:
                text = pytesseract.image_to_string(variant, config=config).strip()
                results.append(text)
            except pytesseract.TesseractError:
                results.append("")

    return max(results, key=score_ocr_text, default="")


def image_to_base64_payload(pil_image, ocr_settings=None):
    cropped_image = get_ocr_source_image(pil_image, ocr_settings)
    max_width = 2200

    if cropped_image.width > max_width:
        ratio = max_width / cropped_image.width
        cropped_image = cropped_image.resize(
            (max_width, int(cropped_image.height * ratio)),
            Image.Resampling.LANCZOS,
        )

    image_buffer = tempfile.SpooledTemporaryFile()
    cropped_image.convert("RGB").save(image_buffer, format="JPEG", quality=92, optimize=True)
    image_buffer.seek(0)
    encoded_image = base64.b64encode(image_buffer.read()).decode("ascii")
    return f"data:image/jpeg;base64,{encoded_image}"


def run_ocr_space_ocr(pil_image, ocr_settings):
    api_key = ocr_settings.get("api_key") or os.environ.get("OCR_SPACE_API_KEY") or "helloworld"
    engine = ocr_settings.get("engine", "3")
    language = ocr_settings.get("language", "auto")

    data = urllib.parse.urlencode(
        {
            "base64Image": image_to_base64_payload(pil_image, ocr_settings),
            "language": language,
            "OCREngine": engine,
            "scale": "true",
            "detectOrientation": "true",
            "isTable": "true" if ocr_settings.get("is_table") else "false",
            "isOverlayRequired": "false",
            "filetype": "JPG",
        }
    ).encode("utf-8")
    request = urllib.request.Request(
        "https://api.ocr.space/parse/image",
        data=data,
        headers={
            "apikey": api_key,
            "Content-Type": "application/x-www-form-urlencoded",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError):
        return ""

    parsed_results = payload.get("ParsedResults") or []
    parsed_text = "\n".join(
        result.get("ParsedText", "").strip()
        for result in parsed_results
        if result.get("ParsedText")
    ).strip()

    return parsed_text


def run_ocr_on_image(pil_image, ocr_settings=None):
    settings = {**DEFAULT_OCR_SETTINGS, **(ocr_settings or {})}

    if settings.get("method") == "ocr_space":
        text = run_ocr_space_ocr(pil_image, settings)
        if text:
            return text

    return run_local_tesseract_ocr(pil_image, settings)


# Function to extract text from each unique slide
def extract_text_from_slides(slides, use_ocr=True, ocr_settings=None):
    slide_texts = []
    progress = st.progress(0)
    
    for idx, slide in enumerate(slides):
        pil_image = Image.fromarray(cv2.cvtColor(slide, cv2.COLOR_BGR2RGB))
        text = ""

        if use_ocr:
            text = run_ocr_on_image(pil_image, ocr_settings) or ""
        
        # Convert to bytes for Streamlit compatibility
        slide_image = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
        pil_image.save(slide_image.name)
        
        slide_texts.append((slide_image.name, text))
        progress.progress((idx + 1) / len(slides))
    
    return slide_texts


def extract_text_from_slide_file(slide_filename, ocr_settings=None):
    try:
        pil_image = Image.open(slide_filename)
        return run_ocr_on_image(pil_image, ocr_settings)
    except OSError:
        return ""


# Function to create a PowerPoint presentation from slides
def create_ppt(slide_texts, use_ocr=True):
    prs = Presentation()
    
    for idx, (slide_filename, text) in enumerate(slide_texts, start=1):
        slide = prs.slides.add_slide(prs.slide_layouts[5])  # Use a blank slide layout
        number_box = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(1.4), Inches(0.35))
        number_frame = number_box.text_frame
        number_frame.clear()
        number_paragraph = number_frame.paragraphs[0]
        number_run = number_paragraph.add_run()
        number_run.text = f"Slide {idx}"
        number_run.font.size = Pt(13)
        number_run.font.bold = True
        number_run.font.color.rgb = RGBColor(70, 70, 70)

        left = top = Inches(1)
        slide.shapes.add_picture(slide_filename, left, top, width=Inches(8.5), height=Inches(6))

        if use_ocr:
            txBox = slide.shapes.add_textbox(Inches(1), Inches(6.5), Inches(8.5), Inches(1))
            tf = txBox.text_frame
            p = tf.add_paragraph()
            p.text = text
    
    # Save PPT to a temporary file
    ppt_temp = tempfile.NamedTemporaryFile(suffix='.pptx', delete=False)
    prs.save(ppt_temp.name)
    return ppt_temp.name


def process_video(
    video_path,
    sample_interval_seconds,
    max_samples,
    max_slides,
    use_ocr,
    ocr_settings,
    start_time_seconds,
    end_time_seconds,
    frame_crop_settings,
):
    with st.spinner("Extracting frames..."):
        frames = extract_frames(
            video_path,
            sample_interval_seconds=sample_interval_seconds,
            max_samples=max_samples,
            start_time_seconds=start_time_seconds,
            end_time_seconds=end_time_seconds,
            frame_crop_settings=frame_crop_settings,
        )

    if not frames:
        st.error("No frames were extracted from the video.")
        return None, None

    with st.spinner("Detecting unique slides..."):
        unique_slides = get_unique_slides(frames)

    if not unique_slides:
        st.error("No unique slides were detected.")
        return None, None

    if len(unique_slides) > max_slides:
        st.warning(
            f"Detected {len(unique_slides)} unique frames. "
            f"Processing the first {max_slides} to keep this run responsive."
        )
        unique_slides = unique_slides[:max_slides]

    if use_ocr:
        with st.spinner("Extracting text from slides..."):
            slide_texts = extract_text_from_slides(
                unique_slides,
                use_ocr=True,
                ocr_settings=ocr_settings,
            )
    else:
        with st.spinner("Preparing slides..."):
            slide_texts = extract_text_from_slides(
                unique_slides,
                use_ocr=False,
                ocr_settings=ocr_settings,
            )

    ppt_filename = create_ppt(slide_texts, use_ocr=use_ocr)
    return slide_texts, ppt_filename


def inject_custom_styles():
    st.markdown(
        """
        <style>
        .block-container {
            padding-top: 2rem;
        }

        .slide-card-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 0.75rem;
            margin-bottom: 0.6rem;
        }

        .slide-number {
            font-weight: 700;
            font-size: 1rem;
        }

        .slide-meta {
            color: rgba(128, 128, 128, 0.95);
            font-size: 0.85rem;
        }

        div[data-testid="stDownloadButton"] button {
            width: 100%;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def copy_text_button(text, button_label, key):
    if not text:
        return

    components.html(
        f"""
        <button id="{key}" style="
            width: 100%;
            border: 1px solid rgba(49, 51, 63, 0.25);
            border-radius: 6px;
            padding: 0.45rem 0.75rem;
            background: transparent;
            color: inherit;
            cursor: pointer;
            font: inherit;
        ">{button_label}</button>
        <script>
        const button = document.getElementById({json.dumps(key)});
        const text = {json.dumps(text)};
        button.addEventListener("click", async () => {{
            try {{
                await navigator.clipboard.writeText(text);
                const originalText = button.innerText;
                button.innerText = "Copied";
                setTimeout(() => button.innerText = originalText, 1200);
            }} catch (error) {{
                button.innerText = "Copy failed";
            }}
        }});
        </script>
        """,
        height=42,
    )


def create_slide_images_zip(slide_texts, selected_indexes=None):
    selected_set = set(selected_indexes) if selected_indexes is not None else None
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zip_file:
        for index, (slide_filename, text) in enumerate(slide_texts):
            if selected_set is not None and index not in selected_set:
                continue

            slide_number = index + 1
            zip_file.write(slide_filename, f"slide_{slide_number:03d}.png")
            if text:
                zip_file.writestr(f"slide_{slide_number:03d}_ocr.txt", text)

    zip_buffer.seek(0)
    return zip_buffer


def store_slide_results(slide_texts, ppt_filename, use_ocr, ocr_settings):
    st.session_state["slide_results"] = {
        "slide_texts": slide_texts,
        "ppt_filename": ppt_filename,
        "use_ocr": use_ocr,
        "ocr_settings": ocr_settings,
    }


def render_processing_controls(prefix):
    with st.container(border=True):
        st.subheader("Processing Settings")
        settings_col_1, settings_col_2, settings_col_3, settings_col_4 = st.columns(4)

        sample_interval_seconds = settings_col_1.slider(
            "Sample every N seconds",
            min_value=5,
            max_value=60,
            value=30,
            step=5,
            key=f"{prefix}_sample_interval_seconds",
        )
        max_samples = settings_col_2.slider(
            "Maximum frames to check",
            min_value=25,
            max_value=500,
            value=120,
            step=25,
            key=f"{prefix}_max_samples",
        )
        max_slides = settings_col_3.slider(
            "Maximum slides to export",
            min_value=25,
            max_value=300,
            value=75,
            step=25,
            key=f"{prefix}_max_slides",
        )
        use_ocr = settings_col_4.toggle(
            "Run OCR during processing",
            value=True,
            help="Turn this off to skip batch text recognition. You can still OCR individual slides later.",
            key=f"{prefix}_use_ocr",
        )

        ocr_settings = DEFAULT_OCR_SETTINGS.copy()
        ocr_col_1, ocr_col_2, ocr_col_3, ocr_col_4 = st.columns([1.2, 1, 1, 1])
        method_label = ocr_col_1.selectbox(
            "OCR method",
            ["Enhanced local Tesseract", "OCR.space cloud"],
            key=f"{prefix}_ocr_method",
            help="Copyfish uses OCR.space engines. Cloud OCR sends the cropped slide image to OCR.space.",
        )
        ocr_settings["method"] = "ocr_space" if method_label == "OCR.space cloud" else "local"

        if ocr_settings["method"] == "ocr_space":
            engine_label = ocr_col_2.selectbox(
                "OCR.space engine",
                ["Engine 3 - best accuracy", "Engine 2 - balanced", "Engine 1 - fastest"],
                key=f"{prefix}_ocr_engine",
            )
            ocr_settings["engine"] = engine_label.split()[1]
            ocr_settings["language"] = ocr_col_3.selectbox(
                "Language",
                ["auto", "eng"],
                key=f"{prefix}_ocr_language",
                help="Auto works with OCR.space engines 2 and 3.",
            )
            ocr_settings["is_table"] = ocr_col_4.checkbox(
                "Table/list mode",
                value=False,
                key=f"{prefix}_ocr_table",
            )
            ocr_settings["api_key"] = st.text_input(
                "OCR.space API key",
                value=os.environ.get("OCR_SPACE_API_KEY", ""),
                type="password",
                key=f"{prefix}_ocr_api_key",
                help="Leave blank to use OCR.space's limited demo key. Add your own key for more reliable results.",
            )
        else:
            ocr_settings["language"] = "eng"

    return sample_interval_seconds, max_samples, max_slides, use_ocr, ocr_settings


def render_video_crop_controls(prefix, video_path, ocr_settings):
    fps, frame_count, duration_seconds = get_video_metadata_from_path(video_path)
    if duration_seconds <= 0:
        st.error("Could not read video duration for crop preview.")
        return 0, None, {"enabled": False}, ocr_settings

    duration_minutes = duration_seconds / 60

    with st.container(border=True):
        st.subheader("Video Range and Crop Preview")
        st.caption(
            f"Loaded video: {duration_minutes:.1f} minutes, "
            f"{frame_count:,} frames, {fps:.2f} fps."
        )

        range_col_1, range_col_2, range_col_3 = st.columns(3)
        start_minutes = range_col_1.number_input(
            "Start time (minutes)",
            min_value=0.0,
            max_value=float(duration_minutes),
            value=0.0,
            step=0.5,
            key=f"{prefix}_start_minutes",
            help="Slides before this timestamp are ignored.",
        )
        end_minutes = range_col_2.number_input(
            "End time (minutes, 0 = end of video)",
            min_value=0.0,
            max_value=float(duration_minutes),
            value=0.0,
            step=0.5,
            key=f"{prefix}_end_minutes",
            help="Use 0 to process until the video ends.",
        )

        preview_max_minutes = end_minutes if end_minutes > start_minutes else duration_minutes
        preview_minutes = range_col_3.slider(
            "Preview frame at minute",
            min_value=float(start_minutes),
            max_value=float(max(start_minutes, preview_max_minutes)),
            value=float(start_minutes),
            step=0.25,
            key=f"{prefix}_preview_minutes",
        )

        frame_crop_enabled = st.checkbox(
            "Crop frame area before slide detection and OCR",
            value=False,
            key=f"{prefix}_frame_crop_enabled",
            help="Use this when the video has a fixed toolbar, presenter camera, black border, or side panel.",
        )
        crop_col_1, crop_col_2, crop_col_3, crop_col_4 = st.columns(4)
        crop_left = crop_col_1.slider(
            "Left %",
            min_value=0,
            max_value=90,
            value=0,
            step=1,
            disabled=not frame_crop_enabled,
            key=f"{prefix}_crop_left",
        )
        crop_right = crop_col_2.slider(
            "Right %",
            min_value=10,
            max_value=100,
            value=100,
            step=1,
            disabled=not frame_crop_enabled,
            key=f"{prefix}_crop_right",
        )
        crop_top = crop_col_3.slider(
            "Top %",
            min_value=0,
            max_value=90,
            value=0,
            step=1,
            disabled=not frame_crop_enabled,
            key=f"{prefix}_crop_top",
        )
        crop_bottom = crop_col_4.slider(
            "Bottom %",
            min_value=10,
            max_value=100,
            value=100,
            step=1,
            disabled=not frame_crop_enabled,
            key=f"{prefix}_crop_bottom",
        )

        if frame_crop_enabled and (crop_left >= crop_right or crop_top >= crop_bottom):
            st.warning("Crop bounds must leave a visible frame area.")

        frame_crop_settings = {
            "enabled": frame_crop_enabled and crop_left < crop_right and crop_top < crop_bottom,
            "left": crop_left,
            "right": crop_right,
            "top": crop_top,
            "bottom": crop_bottom,
        }

        preview_frame = get_video_frame_at_time(video_path, preview_minutes * 60)
        if preview_frame is not None:
            preview_col_1, preview_col_2 = st.columns(2)
            preview_col_1.image(
                cv2.cvtColor(preview_frame, cv2.COLOR_BGR2RGB),
                caption="Original preview frame",
                use_container_width=True,
            )
            cropped_preview = apply_frame_crop(preview_frame, frame_crop_settings)
            preview_col_2.image(
                cv2.cvtColor(cropped_preview, cv2.COLOR_BGR2RGB),
                caption="Frame that will be used for slide detection and OCR",
                use_container_width=True,
            )

        ocr_crop_label = st.selectbox(
            "OCR crop behavior",
            ["Auto-detect slide canvas", "Use generated slide as-is"],
            key=f"{prefix}_ocr_crop_mode",
            help="Use generated slide as-is after manually cropping the frame to the slide area.",
        )
        ocr_settings["crop_mode"] = "manual" if ocr_crop_label == "Use generated slide as-is" else "auto"

    start_time_seconds = start_minutes * 60
    end_time_seconds = end_minutes * 60 if end_minutes > start_minutes else None

    return start_time_seconds, end_time_seconds, frame_crop_settings, ocr_settings


def render_results(slide_texts, ppt_filename, use_ocr, ocr_settings):
    total_slides = len(slide_texts)
    slides_with_text = sum(1 for _, text in slide_texts if text)

    st.success("Processing complete!")

    summary_col_1, summary_col_2, summary_col_3 = st.columns([1, 1, 1.4])
    summary_col_1.metric("Slides generated", total_slides)
    if use_ocr:
        summary_col_2.metric("Slides with OCR text", slides_with_text)
    elif slides_with_text:
        summary_col_2.metric("On-demand OCR text", slides_with_text)
    else:
        summary_col_2.metric("OCR", "Skipped")

    with open(ppt_filename, "rb") as f:
        summary_col_3.download_button(
            "Download PowerPoint Presentation",
            f,
            file_name="presentation.pptx",
            use_container_width=True,
        )

    st.subheader("Generated Slides")
    selection_prefix = f"slide_select_{total_slides}"
    selected_indexes = [
        index for index in range(total_slides)
        if st.session_state.get(f"{selection_prefix}_{index}", False)
    ]

    action_col_1, action_col_2, action_col_3, action_col_4 = st.columns([1, 1, 1.2, 1.6])

    if action_col_1.button("Select all slides", use_container_width=True):
        for index in range(total_slides):
            st.session_state[f"{selection_prefix}_{index}"] = True
        st.rerun()

    if action_col_2.button("Clear selection", use_container_width=True):
        for index in range(total_slides):
            st.session_state[f"{selection_prefix}_{index}"] = False
        st.rerun()

    action_col_3.download_button(
        "Download all slide images",
        create_slide_images_zip(slide_texts),
        file_name="all_slide_images.zip",
        mime="application/zip",
        use_container_width=True,
    )

    if selected_indexes:
        action_col_4.download_button(
            f"Download selected images ({len(selected_indexes)})",
            create_slide_images_zip(slide_texts, selected_indexes),
            file_name="selected_slide_images.zip",
            mime="application/zip",
            use_container_width=True,
        )
    else:
        action_col_4.button(
            "Download selected images",
            disabled=True,
            use_container_width=True,
        )

    for start_index in range(0, total_slides, 2):
        columns = st.columns(2)

        for column_offset, column in enumerate(columns):
            slide_index = start_index + column_offset
            if slide_index >= total_slides:
                continue

            slide_number = slide_index + 1
            slide_filename, text = slide_texts[slide_index]
            if use_ocr:
                text_label = "OCR text found" if text else "No OCR text found"
            else:
                text_label = "OCR text available" if text else "OCR skipped"

            with column:
                with st.container(border=True):
                    st.markdown(
                        f"""
                        <div class="slide-card-header">
                            <span class="slide-number">Slide {slide_number}</span>
                            <span class="slide-meta">{text_label}</span>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    st.checkbox(
                        f"Select Slide {slide_number}",
                        key=f"{selection_prefix}_{slide_index}",
                    )
                    st.image(slide_filename, use_container_width=True)

                    ocr_button_label = (
                        f"Re-run OCR for Slide {slide_number}"
                        if text else f"Run OCR for Slide {slide_number}"
                    )
                    if st.button(
                        ocr_button_label,
                        key=f"run_ocr_slide_{slide_number}",
                        use_container_width=True,
                    ):
                        with st.spinner(f"Running OCR for Slide {slide_number}..."):
                            extracted_text = extract_text_from_slide_file(slide_filename, ocr_settings)

                        if extracted_text is None:
                            st.error("Tesseract OCR is not available on this machine.")
                        else:
                            st.session_state["slide_results"]["slide_texts"][slide_index] = (
                                slide_filename,
                                extracted_text,
                            )
                            updated_slide_texts = st.session_state["slide_results"]["slide_texts"]
                            st.session_state["slide_results"]["ppt_filename"] = create_ppt(
                                updated_slide_texts,
                                use_ocr=use_ocr or any(
                                    item_text for _, item_text in updated_slide_texts
                                ),
                            )
                            st.rerun()

                    if use_ocr or text:
                        with st.expander(f"View OCR text for Slide {slide_number}", expanded=False):
                            if text:
                                st.text(text)
                                copy_text_button(
                                    text,
                                    f"Copy text from Slide {slide_number}",
                                    f"copy_ocr_slide_{slide_number}",
                                )
                            else:
                                st.caption("No readable text was detected for this slide.")


# Streamlit Interface
def main():
    st.set_page_config(page_title="Video Slide Extractor", layout="wide")
    inject_custom_styles()
    st.title("YouTube and Local Video Slide Extractor")
    st.write("Enter a YouTube video link or upload a video file. This app extracts unique slides, runs OCR, and creates a PowerPoint presentation.")

    youtube_tab, upload_tab = st.tabs(["YouTube URL", "Local video upload"])

    with youtube_tab:
        (
            sample_interval_seconds,
            max_samples,
            max_slides,
            use_ocr,
            ocr_settings,
        ) = render_processing_controls("youtube")
        youtube_url = st.text_input("YouTube Video URL", "")

        if st.button("Download video", type="primary"):
            if not youtube_url.strip():
                st.error("Please enter a YouTube video URL.")
                return

            video_temp = tempfile.NamedTemporaryFile(suffix=".mp4", delete=False)
            video_temp.close()
            os.remove(video_temp.name)

            with st.spinner("Downloading video..."):
                video_path = download_video(youtube_url.strip(), output_path=video_temp.name)
                if not video_path:
                    return

            st.session_state["youtube_video"] = {
                "path": video_path,
                "source": youtube_url.strip(),
            }
            st.session_state.pop("slide_results", None)

        youtube_video = st.session_state.get("youtube_video")
        if youtube_video:
            st.success("Video downloaded. Adjust the preview crop/time range, then extract slides.")
            (
                start_time_seconds,
                end_time_seconds,
                frame_crop_settings,
                ocr_settings,
            ) = render_video_crop_controls("youtube", youtube_video["path"], ocr_settings)

            if st.button("Extract slides from downloaded video", type="primary"):
                slide_texts, ppt_filename = process_video(
                    youtube_video["path"],
                    sample_interval_seconds,
                    max_samples,
                    max_slides,
                    use_ocr,
                    ocr_settings,
                    start_time_seconds,
                    end_time_seconds,
                    frame_crop_settings,
                )

                if slide_texts and ppt_filename:
                    store_slide_results(slide_texts, ppt_filename, use_ocr, ocr_settings)

    with upload_tab:
        (
            sample_interval_seconds,
            max_samples,
            max_slides,
            use_ocr,
            ocr_settings,
        ) = render_processing_controls("upload")
        uploaded_file = st.file_uploader(
            "Choose a video from your computer",
            type=["mp4", "mov", "mkv", "avi", "webm"],
        )

        if st.button("Load uploaded video", type="primary"):
            if uploaded_file is None:
                st.error("Please choose a local video file first.")
                return

            extension = os.path.splitext(uploaded_file.name)[1] or ".mp4"
            video_temp = tempfile.NamedTemporaryFile(suffix=extension, delete=False)
            video_temp.close()
            video_path = save_uploaded_video(uploaded_file, video_temp.name)

            st.session_state["upload_video"] = {
                "path": video_path,
                "source": uploaded_file.name,
            }
            st.session_state.pop("slide_results", None)

        upload_video = st.session_state.get("upload_video")
        if upload_video:
            st.success("Video loaded. Adjust the preview crop/time range, then extract slides.")
            (
                start_time_seconds,
                end_time_seconds,
                frame_crop_settings,
                ocr_settings,
            ) = render_video_crop_controls("upload", upload_video["path"], ocr_settings)

            if st.button("Extract slides from loaded video", type="primary"):
                slide_texts, ppt_filename = process_video(
                    upload_video["path"],
                    sample_interval_seconds,
                    max_samples,
                    max_slides,
                    use_ocr,
                    ocr_settings,
                    start_time_seconds,
                    end_time_seconds,
                    frame_crop_settings,
                )

                if slide_texts and ppt_filename:
                    store_slide_results(slide_texts, ppt_filename, use_ocr, ocr_settings)

    if "slide_results" in st.session_state:
        render_results(
            st.session_state["slide_results"]["slide_texts"],
            st.session_state["slide_results"]["ppt_filename"],
            st.session_state["slide_results"]["use_ocr"],
            st.session_state["slide_results"].get("ocr_settings", DEFAULT_OCR_SETTINGS),
        )

if __name__ == "__main__":
    main()
