# Meeting Automation MVP

Local-first meeting intelligence app for turning audio, transcripts, and supporting materials into investor-style outputs:

- Clean transcript preview and download
- Interview memo markdown for Obsidian / knowledge base workflows
- Weekly report PPT generation with optional enhancement review
- OI news report draft generation

The project is designed to run locally first, while keeping the provider layer environment-driven so it can later move to a cloud VM or object-storage-backed deployment.

## Current Workflow

1. Add audio or an existing transcript.
2. Generate a cleaned transcript.
3. Choose a memo style and generate an interview memo.
4. Optionally add supporting materials and generate a weekly report PPT from the memo.
5. Download outputs or save markdown to an Obsidian vault.

## Project Structure

```text
meeting-automation-mvp/
  assets/       PPT assets and reusable slide templates
  docs/         workflow and template documentation
  public/       browser UI
  scripts/      local ASR and PPT helper scripts
  src/          Express server, pipeline, providers, storage
  templates/    JSON output templates
```

Runtime files are written under `data/` and ignored by git.

## Run Locally

```powershell
npm install
Copy-Item .env.example .env
npm start
```

Open [http://localhost:3100](http://localhost:3100).

## Configuration

Set provider keys in your shell or in a local `.env` file. Do not commit real keys.

### LLM

- `ALI_API_KEY`: Alibaba / Qwen-compatible chat completions
- `ALI_MODEL`: default `qwen-plus`
- `ALI_BASE_URL`: default `https://dashscope.aliyuncs.com/compatible-mode/v1`
- `OPENAI_API_KEY`: optional OpenAI provider
- `OPENAI_MODEL`: optional OpenAI model

The UI supports `auto`, `ali`, `openai`, and `none`. `auto` prefers OpenAI when configured, then Alibaba, then rule-based fallback.

### ASR

- `ALI_API_KEY`: also used by Alibaba Qwen ASR
- `ALI_ASR_MODEL`: default `qwen3-asr-flash`
- `WHISPER_API_KEY`: optional OpenAI-compatible Whisper endpoint
- `LOCAL_WHISPER_MODEL`: optional local Whisper model name

Audio can be chunked/compressed for longer recordings through the Alibaba ASR path.

### Research Enrichment

Optional public-source enrichment can be enabled with:

- `TAVILY_API_KEY`
- `SERPAPI_API_KEY`

The weekly-report flow surfaces proposed enrichment in a review table so user-provided material stays distinguishable from public-source additions.

## Tests

```powershell
npm test
```

## Notes For GitHub

- `data/`, `generated/`, `node_modules/`, local env files, and cache files are ignored.
- Keep real meeting transcripts, company PDFs, generated PPTs, and API keys out of the repository.
- Commit only source code, reusable templates, docs, and non-confidential assets.
