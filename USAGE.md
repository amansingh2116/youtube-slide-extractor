# Usage Guide

This guide explains how to use the YouTube and Local Video Slide Extractor and what each processing option does.

## Start the App

Run:

```bash
streamlit run app.py
```

Open the local URL Streamlit prints in the terminal.

## Input Modes and Video Queue

The app can load up to five videos at a time. Videos can come from YouTube, local uploads, or a mix of both.

After loading videos, use **Loaded Videos** to choose which video to work on. Each selected video gets its own processing settings, crop preview, OCR settings, and generated results.

### YouTube URL

Use this tab when the source video is on YouTube.

1. Open the **YouTube URL** tab.
2. Paste one or more complete YouTube URLs, one per line.
3. Click **Download YouTube video(s)**.
4. The downloaded videos are added to **Loaded Videos**.
5. Choose a loaded video and process it.

The app downloads each video with `yt-dlp`, saves it temporarily, and adds it to the queue.

### Local Video Upload

Use this tab when the video file is already on your computer.

1. Open the **Local video upload** tab.
2. Choose one or more video files from your computer.
3. Click **Load uploaded video(s)**.
4. The uploaded videos are added to **Loaded Videos**.
5. Choose a loaded video and process it.

Supported upload extensions are `mp4`, `mov`, `mkv`, `avi`, and `webm`.

### Loaded Videos

The **Loaded Videos** section appears after at least one video is loaded.

Use it to:

- choose the active video
- remove the selected video from the queue
- see how many of the five slots are currently used

To process multiple videos, choose the first video, configure its settings, extract slides, then choose the next video and repeat.

## Processing Settings

Each loaded video has its own settings. Switching between loaded videos keeps those settings separate.

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

## Advanced Crop and Time Range

These controls appear after you choose a loaded video. They use a real frame from that selected video, so you can visually choose the crop area before extraction starts.

### Start and End Time

Use these controls to process only part of a video.

- **Start time (minutes)** ignores everything before that timestamp.
- **End time (minutes, 0 = end of video)** stops processing after that timestamp.

This is useful when:

- The lecture has a long intro.
- You only need one topic/chapter.
- You want to test settings on a short section before processing the full video.

Example:

- Start time: `10`
- End time: `25`

The app only samples frames between minute 10 and minute 25. The preview frame slider also follows this selected range.

### Crop Frame Area Before Slide Detection and OCR

Enable this when the slide is always in the same area of the video frame.

The crop is applied before:

- duplicate slide detection
- generated slide gallery
- OCR
- PowerPoint export

Use the percentage sliders:

- **Left %** removes content from the left edge.
- **Right %** keeps content up to that percentage from the left.
- **Top %** removes content from the top edge.
- **Bottom %** keeps content up to that percentage from the top.

For example, if a presenter camera is always on the right side, reduce **Right %** until the presenter camera is outside the cropped frame.

The app shows:

- the original preview frame
- the cropped frame that will be used for slide detection and OCR

Adjust the crop sliders until the cropped preview contains the slide area you want.

This is useful for videos with:

- PDF/browser toolbars
- Windows or macOS taskbars
- presenter camera bubbles
- black borders
- side panels
- chat windows
- fixed watermarks or overlays

### OCR Crop Behavior

This controls what image OCR receives.

Use **Auto-detect slide canvas** when:

- the full sampled frame still contains toolbar/taskbar/presenter noise
- you did not manually crop the frame
- the slide is surrounded by screen-recording UI

Use **Use generated slide as-is** when:

- you already used the manual frame crop to isolate the slide
- auto-crop cuts off useful slide content
- the slide background is not white
- the video uses a non-standard slide layout

This setting is important because not every video has the same layout. Auto-crop is helpful for many lecture recordings, but manual crop plus generated-slide OCR is more reliable when the layout is unusual.

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

If auto-crop does not work for a specific video, enable manual frame cropping and set OCR crop behavior to **Use generated slide as-is**.

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
- Slide selection checkbox.
- Slide image.
- OCR status.
- **Run OCR** or **Re-run OCR** button.
- Expandable OCR text area when text exists.
- Copy button for OCR text.

### Download Slide Images

The results gallery includes image export controls above the slides.

Use **Download all slide images** to download every generated slide image as a ZIP file.

Use the slide checkboxes to select only the slides you want, then click **Download selected images**.

The ZIP files include:

- `slide_001.png`, `slide_002.png`, etc.
- `slide_001_ocr.txt`, etc. when that slide has OCR text.

Use **Select all slides** to mark every slide, or **Clear selection** to reset the checkboxes.

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
2. Load the video and use the preview frame controls.
3. Enable **Crop frame area before slide detection and OCR** and crop out fixed UI noise.
4. Set **OCR crop behavior** to **Use generated slide as-is**.
5. Re-run OCR for that slide.
6. Switch OCR method to **OCR.space cloud**.
7. Use **Engine 3 - best accuracy**.
8. Enable **Table/list mode** for contents pages or tabular slides.
9. Use a personal OCR.space API key for more reliable cloud results.

### Auto-Crop Cuts Off Slide Content

Switch **OCR crop behavior** to **Use generated slide as-is**.

If the generated slide still includes too much noise, enable manual frame crop and adjust the crop percentages until the generated slide itself looks clean.

### Only Need Part of the Video

Use **Start time** and **End time** in **Advanced crop and time range**. This reduces processing time and avoids extracting slides from irrelevant sections.

### Local OCR Does Not Work

Make sure Tesseract is installed and available on your system.

On Windows, the app automatically checks:

```text
C:\Program Files\Tesseract-OCR\tesseract.exe
```

If Tesseract is installed somewhere else, add it to your system `PATH`.

### YouTube Download Fails

Make sure the URL is complete and includes a valid YouTube video ID. Some videos may fail if they are private, region-blocked, age-restricted, or unavailable to `yt-dlp`.
