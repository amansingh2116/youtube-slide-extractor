# YouTube and Local Video Slide Extractor

This Streamlit app converts presentation-style videos into a numbered slide gallery and downloadable PowerPoint file. It supports a queue of up to five YouTube or local videos, samples frames from each video, removes near-duplicate slides, optionally runs OCR, and lets you OCR or re-OCR individual slides after processing.

## Features

- **Two input modes**: load YouTube URLs or upload local video files.
- **Five-video queue**: load up to five videos and process them one after another.
- **Per-video settings**: each loaded video has independent processing, crop, OCR, and result controls.
- **Efficient frame sampling**: checks frames at configurable time intervals instead of reading every frame.
- **Video time range selection**: process only a selected part of a video.
- **Manual frame-area crop**: crop every sampled frame before slide detection and OCR, useful for fixed toolbars, camera overlays, borders, or side panels.
- **Duplicate slide filtering**: uses perceptual hashing to keep visually distinct slide frames.
- **Numbered slide gallery**: generated slides are displayed as numbered cards.
- **Slide image export**: select individual slides or download all generated slide images as a ZIP.
- **Optional OCR**: run OCR during processing or skip it for faster slide extraction.
- **Per-slide OCR**: run or re-run OCR for any individual slide from the results gallery.
- **Copy OCR text**: copy recognized text for each slide that has OCR output.
- **Improved OCR preprocessing**: crops to the likely slide canvas, removes screen-recording noise, upscales the image, and tries multiple Tesseract modes.
- **OCR.space support**: optional cloud OCR mode inspired by Copyfish/OCR.space workflows.
- **PowerPoint export**: downloads a `.pptx` with slide numbers and OCR text when available.

## Tech Stack

- **Streamlit** for the web UI.
- **OpenCV** for video reading and image processing.
- **Pillow** for image conversion and saving.
- **imagehash** for perceptual duplicate detection.
- **pytesseract** for local OCR.
- **OCR.space API** for optional cloud OCR.
- **yt-dlp** for YouTube downloads.
- **python-pptx** for PowerPoint generation.

## Installation

1. Clone the repository:

```bash
git clone git@github.com:Adityakeshav/youtube-slide-extractor.git
cd youtube-slide-extractor
```

2. Install Python dependencies:

```bash
pip install -r requirements.txt
```

3. Install Tesseract OCR for local OCR:

- Windows: install Tesseract OCR and make sure it is available at `C:\Program Files\Tesseract-OCR\tesseract.exe` or on your `PATH`.
- Linux:

```bash
sudo apt install tesseract-ocr
```

- macOS:

```bash
brew install tesseract
```

4. Optional: set an OCR.space API key for cloud OCR:

```bash
set OCR_SPACE_API_KEY=your_api_key_here
```

On macOS/Linux:

```bash
export OCR_SPACE_API_KEY=your_api_key_here
```

You can also paste the key directly into the app when using OCR.space cloud mode.

## Running the App

```bash
streamlit run app.py
```

Then open the local Streamlit URL shown in the terminal.

## Quick Usage

1. Choose **YouTube URL** or **Local video upload**.
2. Load one or more videos. The queue supports up to five videos total.
3. Choose a video from **Loaded Videos**.
4. Adjust processing settings for the selected video.
5. Choose whether to run OCR during processing.
6. Choose an OCR method:
   - **Enhanced local Tesseract** for offline OCR.
   - **OCR.space cloud** for optional cloud OCR engines.
7. Use the preview frame to choose the video time range and visual crop.
8. Extract slides from the selected video.
9. Review the generated numbered slide gallery.
10. Select slides and download selected/all slide images if needed.
11. Run or re-run OCR for any slide if needed.
12. Copy OCR text from individual slides.
13. Download the generated PowerPoint.
14. Select another loaded video and repeat.

See [USAGE.md](USAGE.md) for the full workflow and setting details.

## How It Works

The app first downloads or saves each input video to a persistent temporary file and adds it to the loaded-video queue. After you choose a video from the queue, it reads video metadata such as FPS, frame count, and duration, then shows a real preview frame. You can use that preview to choose the start/end time range and crop the frame area visually before slide extraction begins.

When extraction starts, the app samples frames by timestamp. The number of sampled frames is controlled by **Sample every N seconds**, **Maximum frames to check**, and the optional start/end time range.

If manual frame cropping is enabled, each sampled frame is cropped before duplicate detection. This means slide detection, OCR, gallery display, and PowerPoint export all use the cropped frame. After sampling and optional cropping, each frame is converted to a perceptual hash. Frames that are too visually similar to previous frames are skipped, leaving only unique slide-like frames.

When OCR is enabled, the app can either auto-detect the likely presentation canvas before OCR or use the generated slide image as-is. Auto-detect helps avoid toolbar text, taskbar text, presenter bubbles, and other screen-recording noise. If you manually crop the video frame to the slide area, use the generated-slide OCR mode so OCR does not crop again. Local OCR uses Tesseract with multiple preprocessing and page-segmentation modes, then keeps the result with the strongest text score. Cloud OCR uses OCR.space when selected.

The final results are stored by video in Streamlit session state, so you can process one video, switch to another, then return to the first video's gallery. A PowerPoint file is generated from each video's slide images and OCR text.
