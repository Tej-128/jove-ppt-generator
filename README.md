# JoVE PPT Generator

A Streamlit-based PowerPoint generator for creating JoVE-style biology lecture slide decks from lesson files.

The app takes a ZIP folder containing JoVE lesson materials, reads the lesson PageText and Transcript files, uses AI to generate structured slide content, selects relevant visuals from lesson MP4 videos, and outputs a downloadable `.pptx` presentation.

## What This App Does

* Generates JoVE-style lecture decks from uploaded lesson ZIP files
* Supports PageText, Transcript, and MP4 lesson inputs
* Uses OpenAI for slide content generation
* Uses lesson MP4 videos as the first-priority image source
* Extracts relevant video frames based on slide content and transcript context
* Removes or crops JoVE watermark areas from extracted video frames where possible
* Uses AI-generated fallback images only when no suitable JoVE image is available
* Builds PowerPoint slides using JoVE layout, color, and formatting rules
* Keeps the cover first, adds a chapter definition on the cover, and adds a chapter overview slide before lesson-specific slides
* Uses strategic discussion Question/Answer pairs only where useful, with at least one pair and no more than six pairs per deck
* Adds short bottom narration transition captions to normal concept slides where natural
* Adds 1-5 word figure legends below inserted slide images, excluding table Image-column cells
* Adds slide numbers and copyright footer automatically
* Generates a QA report with slide/image details

## Required File Structure

Upload a ZIP containing lesson files.

Supported naming examples:

```text
10661_Pagetext.docx
10661_Transcript.docx
10661_video.mp4
```

Older naming formats are also supported:

```text
10661_LessonName_Pagetext.docx
10661_LessonName_Transcript.docx
10661_LessonName.mp4
```

Each lesson should have:

```text
Lesson ID + PageText DOCX
Lesson ID + Transcript DOCX
Lesson ID + MP4 video
```

The MP4 file is required for JoVE image extraction.

## GitHub Files

The repo should include:

```text
app.py
pipeline.py
ai_generator.py
planner.py
ppt_builder.py
image_sourcing.py
video_sourcing.py
style_guide.py
requirements.txt
jove_logo.png
.streamlit/config.toml
```

## Streamlit Secrets

In Streamlit Cloud, add the following secret:

```toml
OPENAI_API_KEY = "your-openai-api-key-here"
```

The app currently uses OpenAI for content generation, vision-based frame selection, and AI fallback images.

## Streamlit Config

The `.streamlit/config.toml` file should contain:

```toml
[server]
maxUploadSize = 1500
maxMessageSize = 1500
```

This allows larger ZIP uploads and larger app responses.

## Current Image Rules

Image sourcing follows this priority:

```text
1. JoVE video frame from the lesson MP4
2. Clean/inpaint watermark if possible
3. Crop watermark area if cleanup looks bad
4. Choose another JoVE frame if needed
5. Use AI-generated image fallback only if no suitable JoVE frame is available
```

All AI fallback usage should be treated as an exception and reviewed in the QA report.

## Slide Formatting Rules

The generated decks follow the JoVE presentation guideline as closely as possible:

* 20 x 11.25 inch widescreen slides
* White background
* JoVE blue table headers
* Text left, image right for normal concept slides
* Full-width tables for table slides
* Images embedded inside table rows where possible
* No separate image beside table slides
* JoVE logo on the slide itself
* No JoVE watermark inside embedded images where possible
* No writer name, author name, reviewer name, or prepared-by text
* Bottom-center copyright footer
* Bottom-right slide numbering

## Copyright Footer

Each slide uses:

```text
Copyright © 2026 MyJoVE Corporation. All rights reserved
```

## Running Locally

Install dependencies:

```bash
pip install -r requirements.txt
```

Run the app:

```bash
streamlit run app.py
```

## Deployment

This app is designed for Streamlit Cloud.

After updating GitHub files:

1. Commit the changes to the main branch.
2. Go to Streamlit Cloud.
3. Open the app settings.
4. Click **Reboot app**.
5. Test with a small ZIP before uploading a large chapter ZIP.

## Notes

This generator is still under active improvement. The current version includes guideline hardening for formatting, table layout, watermark handling, and image sourcing, but generated decks should still be reviewed for scientific accuracy, image relevance, table fit, and formatting quality before final use.
