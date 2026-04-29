# PowerPoint Notes Summarizer

A minimal local web app for extracting PowerPoint speaker notes, summarizing them with an OpenAI model, comparing before/after notes, and downloading a copy of the deck with summarized notes.

## Run

```bash
python3 app.py
```

The server prints the local URL it is using. By default it starts at port `8000` and automatically tries nearby ports if that one is busy.

## Inference Providers

The app supports three inference paths:

- OpenAI Codex login using local OAuth through `codex exec`.
- OpenAI API key using the Responses API.
- OpenRouter API key using OpenRouter chat completions.

You can enter API keys in the app when needed, or set them before launch:

```bash
export OPENAI_API_KEY="your-openai-key"
export OPENROUTER_API_KEY="your-openrouter-key"
```

For Codex login, check local auth with:

```bash
codex login status
```

## What It Does

- Accepts `.pptx` files directly.
- Accepts `.ppt` files when LibreOffice or `soffice` is installed, converting them to `.pptx` before processing.
- Reads the presentation slide order.
- Finds speaker notes attached to each slide.
- Lets you choose Codex OAuth, OpenAI API key, or OpenRouter API key.
- Lets you edit the summarization prompt and model before running.
- Writes summarized notes back into the deck.
- Returns a `.pptx` download.

## Default Prompt

```text
Transform the provided detailed slide notes into concise bullet-point reminders for live presentation delivery.

INSTRUCTIONS:

- Convert verbose explanations into short, memorable bullet points
- Focus on key topics, concepts, and transitions the presenter needs to remember
- Preserve technical terms, product names, and specific numbers/statistics
- Keep the original slide numbering and structure
- Include any instructor notes or presenter directions in [brackets]
- Maintain logical flow between points
- Remove filler words and redundant explanations


BULLET POINT GUIDELINES:


- Keep the titles for each slide as "## Slide [number]" without any further text after the number, followed by the summarized slide notes in the next line
- DO NOT use any formatting in the output text, this text will be copied and pasted in a .txt file, avoid rich text. Do bullet point with a "- "
- Maximum 5-10 words per bullet when possible
- Use action verbs and keywords as memory triggers
- Include specific examples, names, or statistics that must be mentioned
- Note key transitions between sections
- Highlight any demonstrations, clicks, or interactive elements
- Keep audience engagement cues (questions, discussion points)

GOAL: Create speaking notes that allow the presenter to glance quickly and speak naturally while covering all essential content without reading verbatim.
```

## Tests

```bash
python3 -m unittest discover -s tests
```
