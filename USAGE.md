# Usage Guide

This guide explains how to use the YouTube and Local Video Slide Extractor and what each processing option does.

## Start the App

Run:

```bash
streamlit run app.py
```

Open the local URL Streamlit prints in the terminal.

## Input Modes

### YouTube URL

Use this tab when the source video is on YouTube.

1. Open the **YouTube URL** tab.
2. Configure the processing settings on the page.
3. Paste a complete YouTube URL.
4. Click **Download and extract slides**.

The app downloads the video with `yt-dlp`, saves it temporarily, and processes that file.

### Local Video Upload

Use this tab when the video file is already on your computer.

1. Open the **Local video upload** tab.
2. Configure the processing settings on the page.
3. Choose a video file from your computer.
4. Click **Extract slides from upload**.

Supported upload extensions are `mp4`, `mov`, `mkv`, `avi`, and `webm`.

## Processing Settings

Each input tab has its own settings.

### Sample Every N Seconds

Controls how often the app checks the video for a slide frame.

- Lower values check more frames and can catch quick slide changes.
- Higher values are faster and better for long lectures with slower slide changes.
- A good default for long lectures is `30` seconds.

### Maximum Frames to Check

Caps how many sampled frames the app will inspect.

- Higher values may find more slides.
- Lower values finish faster.
- For long lectures, `120` is a practical starting point.

### Maximum Slides to Export

Caps how many unique slides are kept for OCR and PowerPoint export.

This prevents very long videos from creating huge result sets.

### Run OCR During Processing

When enabled, OCR runs on every detected unique slide during the main processing step.

When disabled, the app extracts the slides faster. You can still run OCR later on individual slides from the generated gallery.

## OCR Methods

### Enhanced Local Tesseract

This mode runs OCR locally on your computer with Tesseract.

Before sending an image to Tesseract, the app:

1. Detects the likely white slide canvas.
2. Crops away screen-recording noise such as toolbars, taskbars, and presenter bubbles.
3. Upscales the crop.
4. Denoises and enhances contrast.
5. Runs multiple Tesseract page-segmentation modes.
6. Scores the OCR outputs and keeps the best-looking result.

This mode works offline once Tesseract is installed.

### OCR.space Cloud

This mode sends the cropped slide image to OCR.space.

It is useful when local Tesseract struggles with a slide. The app exposes OCR.space engine choices similar to Copyfish-style workflows:

- **Engine 3 - best accuracy**: recommended first for difficult slides.
- **Engine 2 - balanced**: useful fallback.
- **Engine 1 - fastest**: fastest, but usually less accurate.

You can leave the API key blank to use OCR.space's limited demo key, but a personal API key is more reliable.

### Table/List Mode

Enable this when the slide contains contents pages, tables, ledgers, item lists, or structured rows.

## Results Gallery

After processing, the app shows a generated slide gallery.

Each slide card includes:

- Slide number.
- Slide image.
- OCR status.
- **Run OCR** or **Re-run OCR** button.
- Expandable OCR text area when text exists.
- Copy button for OCR text.

### Per-Slide OCR

Use this when:

- You skipped batch OCR.
- One slide has poor OCR.
- You changed OCR settings and want to retry a specific slide.

Click **Run OCR for Slide X** or **Re-run OCR for Slide X**. The app updates that slide's OCR text and regenerates the PowerPoint download so it includes the latest OCR text.

### Copy OCR Text

When OCR text exists for a slide, expand the OCR section and click the copy button. The text is copied to your clipboard in the browser.

## PowerPoint Export

The app generates a `.pptx` file with:

- One generated slide per detected unique video slide.
- A visible slide number.
- The slide image.
- OCR text when available.

Click **Download PowerPoint Presentation** in the results summary.

## Troubleshooting

### Processing Takes Too Long

Increase **Sample every N seconds** or lower **Maximum frames to check**.

For long lectures, try:

- Sample every N seconds: `30` or `45`
- Maximum frames to check: `100` to `150`
- Maximum slides to export: `75`

### Too Many Duplicate Slides

Increase **Sample every N seconds** so fewer near-identical frames are checked.

The app also removes duplicates with perceptual hashing, but sampling fewer frames can make long videos more manageable.

### Missing Slides

Decrease **Sample every N seconds** so the app checks the video more frequently.

For fast-moving lectures, try `5` or `10` seconds.

### OCR Is Poor

Try these in order:

1. Click **Re-run OCR** for that specific slide.
2. Switch OCR method to **OCR.space cloud**.
3. Use **Engine 3 - best accuracy**.
4. Enable **Table/list mode** for contents pages or tabular slides.
5. Use a personal OCR.space API key for more reliable cloud results.

### Local OCR Does Not Work

Make sure Tesseract is installed and available on your system.

On Windows, the app automatically checks:

```text
C:\Program Files\Tesseract-OCR\tesseract.exe
```

If Tesseract is installed somewhere else, add it to your system `PATH`.

### YouTube Download Fails

Make sure the URL is complete and includes a valid YouTube video ID. Some videos may fail if they are private, region-blocked, age-restricted, or unavailable to `yt-dlp`.
