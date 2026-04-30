# Chinese ASR Choice: SenseVoice / FunASR

## Recommended default

For the Chinese ASR track of this MVP, the best practical open-source choice is:

- `SenseVoice-Small` served through `FunASR / SenseVoice FastAPI`

## Why this is the best fit

- Strong Chinese and Cantonese recognition
- Better Chinese-oriented recognition than Whisper in the official benchmark summary
- Very fast inference
- Open-source model and service path
- Supports long audio with VAD-based segmentation
- Easy to keep private by self-hosting

## Official references

- FunASR GitHub: https://github.com/modelscope/FunASR
- SenseVoice GitHub: https://github.com/FunAudioLLM/SenseVoice
- FunASR homepage: https://funaudiollm.github.io/funasr/

## Evidence from official docs

FunASR lists:

- `SenseVoiceSmall` with ASR, ITN, LID, SER, AED support
- `paraformer-zh` as a Mandarin-only option
- `Fun-ASR-Nano` as a newer broader speech recognition model

For this MVP, `SenseVoice-Small` is the most balanced choice because it is lighter and more deployment-ready than a newest large research release, while being more Chinese-friendly than Whisper.

## Why not choose FireRedASR first

FireRedASR2 is very strong in recent research results, but for this MVP it is not the best first integration target because:

- it is newer and heavier
- deployment patterns are less standard for a quick MVP path
- SenseVoice already has a clearer service deployment story for practical integration

So the recommendation is:

- `Chinese ASR`: SenseVoice / FunASR
- `Fallback / multilingual`: Whisper-compatible API

## Official deployment references

SenseVoice official repo shows:

- FastAPI deployment
- Docker build
- Docker run with CPU or GPU

Official examples:

```bash
export SENSEVOICE_DEVICE=cuda:0
fastapi run --port 50000
```

```bash
docker build -t sensevoice .
docker run --gpus all -p 50000:50000 sensevoice
```

```bash
docker run -e SENSEVOICE_DEVICE=cpu -p 50000:50000 sensevoice
```

## App integration contract

This project now expects a SenseVoice-compatible HTTP endpoint via:

- `SENSEVOICE_API_URL`

Optional:

- `SENSEVOICE_API_KEY`
- `SENSEVOICE_LANGUAGE`
- `SENSEVOICE_USE_ITN`

The current parser accepts common response shapes such as:

```json
{ "text": "..." }
```

```json
{ "result": [{ "text": "..." }] }
```

```json
{ "result": { "text": "..." } }
```

## Suggested production layout

For cloud deployment:

1. Run the Node app as the API and orchestration layer
2. Run SenseVoice as a dedicated internal ASR service
3. Run Whisper as another service or external API
4. Keep provider selection at the job level
