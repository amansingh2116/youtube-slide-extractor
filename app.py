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
import yt_dlp as youtube_dl
from pptx import Presentation
from pptx.util import Inches
from pptx.util import Pt
from pptx.dml.color import RGBColor

TESSERACT_WINDOWS_PATH = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


if os.path.exists(TESSERACT_WINDOWS_PATH):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_WINDOWS_PATH


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


# Function to extract frames from the video at specific intervals
def extract_frames(video_path, sample_interval_seconds=5, max_samples=400):
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        st.error("Could not open the video file. Please try another video.")
        return []

    fps, frame_count, duration_seconds = get_video_metadata(cap)
    frames = []

    if fps <= 0 or frame_count <= 0:
        frame_interval = 150
        frame_id = 0

        while cap.isOpened() and len(frames) < max_samples:
            ret, frame = cap.read()
            if not ret:
                break
            if frame_id % frame_interval == 0:
                frames.append(frame)
            frame_id += 1

        cap.release()
        return frames

    frame_step = max(1, int(fps * sample_interval_seconds))
    sample_count = min(max_samples, math.ceil(duration_seconds / sample_interval_seconds))
    st.caption(
        f"Video length: {duration_seconds / 60:.1f} minutes. "
        f"Checking {sample_count} sampled frame(s)."
    )
    progress = st.progress(0)

    for sample_index in range(sample_count):
        target_msec = sample_index * sample_interval_seconds * 1000
        target_frame = min(sample_index * frame_step, frame_count - 1)

        cap.set(cv2.CAP_PROP_POS_MSEC, target_msec)
        ret, frame = cap.read()

        if not ret:
            cap.set(cv2.CAP_PROP_POS_FRAMES, target_frame)
            ret, frame = cap.read()

        if ret:
            frames.append(frame)

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

# Function to extract text from each unique slide
def extract_text_from_slides(slides):
    slide_texts = []
    tesseract_available = shutil.which("tesseract") or os.path.exists(TESSERACT_WINDOWS_PATH)
    progress = st.progress(0)
    
    for idx, slide in enumerate(slides):
        pil_image = Image.fromarray(cv2.cvtColor(slide, cv2.COLOR_BGR2RGB))
        text = ""

        if tesseract_available:
            try:
                text = pytesseract.image_to_string(pil_image).strip()
            except pytesseract.TesseractError:
                text = ""
        
        # Convert to bytes for Streamlit compatibility
        slide_image = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
        pil_image.save(slide_image.name)
        
        slide_texts.append((slide_image.name, text))
        progress.progress((idx + 1) / len(slides))
    
    return slide_texts

# Function to create a PowerPoint presentation from slides
def create_ppt(slide_texts):
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
        txBox = slide.shapes.add_textbox(Inches(1), Inches(6.5), Inches(8.5), Inches(1))
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = text
    
    # Save PPT to a temporary file
    ppt_temp = tempfile.NamedTemporaryFile(suffix='.pptx', delete=False)
    prs.save(ppt_temp.name)
    return ppt_temp.name


def process_video(video_path, sample_interval_seconds, max_samples, max_slides):
    with st.spinner("Extracting frames..."):
        frames = extract_frames(
            video_path,
            sample_interval_seconds=sample_interval_seconds,
            max_samples=max_samples,
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

    with st.spinner("Extracting text from slides..."):
        slide_texts = extract_text_from_slides(unique_slides)

    ppt_filename = create_ppt(slide_texts)
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


def render_results(slide_texts, ppt_filename):
    total_slides = len(slide_texts)
    slides_with_text = sum(1 for _, text in slide_texts if text)

    st.success("Processing complete!")

    summary_col_1, summary_col_2, summary_col_3 = st.columns([1, 1, 1.4])
    summary_col_1.metric("Slides generated", total_slides)
    summary_col_2.metric("Slides with OCR text", slides_with_text)

    with open(ppt_filename, "rb") as f:
        summary_col_3.download_button(
            "Download PowerPoint Presentation",
            f,
            file_name="presentation.pptx",
            use_container_width=True,
        )

    st.subheader("Generated Slides")

    for start_index in range(0, total_slides, 2):
        columns = st.columns(2)

        for column_offset, column in enumerate(columns):
            slide_index = start_index + column_offset
            if slide_index >= total_slides:
                continue

            slide_number = slide_index + 1
            slide_filename, text = slide_texts[slide_index]
            text_label = "OCR text found" if text else "No OCR text found"

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
                    st.image(slide_filename, use_container_width=True)

                    with st.expander(f"View OCR text for Slide {slide_number}", expanded=False):
                        if text:
                            st.text(text)
                        else:
                            st.caption("No readable text was detected for this slide.")


# Streamlit Interface
def main():
    st.set_page_config(page_title="Video Slide Extractor", layout="wide")
    inject_custom_styles()
    st.title("YouTube and Local Video Slide Extractor")
    st.write("Enter a YouTube video link or upload a video file. This app extracts unique slides, runs OCR, and creates a PowerPoint presentation.")

    with st.sidebar:
        st.header("Processing")
        sample_interval_seconds = st.slider(
            "Sample every N seconds",
            min_value=5,
            max_value=60,
            value=30,
            step=5,
        )
        max_samples = st.slider(
            "Maximum frames to check",
            min_value=25,
            max_value=500,
            value=120,
            step=25,
        )
        max_slides = st.slider(
            "Maximum slides for OCR/PPT",
            min_value=25,
            max_value=300,
            value=75,
            step=25,
        )

    youtube_tab, upload_tab = st.tabs(["YouTube URL", "Local video upload"])

    with youtube_tab:
        youtube_url = st.text_input("YouTube Video URL", "")

        if st.button("Download and extract slides", type="primary"):
            if not youtube_url.strip():
                st.error("Please enter a YouTube video URL.")
                return

            with tempfile.TemporaryDirectory() as temp_dir:
                temp_video_path = os.path.join(temp_dir, "video.mp4")

                with st.spinner("Downloading video..."):
                    video_path = download_video(youtube_url.strip(), output_path=temp_video_path)
                    if not video_path:
                        return

                slide_texts, ppt_filename = process_video(
                    video_path,
                    sample_interval_seconds,
                    max_samples,
                    max_slides,
                )

            if slide_texts and ppt_filename:
                render_results(slide_texts, ppt_filename)

    with upload_tab:
        uploaded_file = st.file_uploader(
            "Choose a video from your computer",
            type=["mp4", "mov", "mkv", "avi", "webm"],
        )

        if st.button("Extract slides from upload", type="primary"):
            if uploaded_file is None:
                st.error("Please choose a local video file first.")
                return

            with tempfile.TemporaryDirectory() as temp_dir:
                extension = os.path.splitext(uploaded_file.name)[1] or ".mp4"
                temp_video_path = os.path.join(temp_dir, f"uploaded_video{extension}")
                video_path = save_uploaded_video(uploaded_file, temp_video_path)

                slide_texts, ppt_filename = process_video(
                    video_path,
                    sample_interval_seconds,
                    max_samples,
                    max_slides,
                )

            if slide_texts and ppt_filename:
                render_results(slide_texts, ppt_filename)

if __name__ == "__main__":
    main()
