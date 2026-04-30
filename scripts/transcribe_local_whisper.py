import json
import sys

from faster_whisper import WhisperModel

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")


def main():
    if len(sys.argv) < 2:
        raise SystemExit("Usage: transcribe_local_whisper.py <audio_path> [options_json_path]")

    audio_path = sys.argv[1]
    options_path = sys.argv[2] if len(sys.argv) > 2 else None
    options = {}
    if options_path:
        with open(options_path, "r", encoding="utf-8") as fh:
            options = json.load(fh)

    model_name = options.get("model", "small")
    device = options.get("device", "cpu")
    compute_type = options.get("compute_type", "int8")
    language = options.get("language") or None
    initial_prompt = options.get("initial_prompt") or None
    beam_size = int(options.get("beam_size", 5))
    best_of = int(options.get("best_of", 5))
    temperature = float(options.get("temperature", 0))
    condition_on_previous_text = bool(options.get("condition_on_previous_text", False))
    vad_filter = bool(options.get("vad_filter", True))

    model = WhisperModel(model_name, device=device, compute_type=compute_type)
    segments_iter, info = model.transcribe(
        audio_path,
        language=language,
        initial_prompt=initial_prompt,
        beam_size=beam_size,
        best_of=best_of,
        temperature=temperature,
        condition_on_previous_text=condition_on_previous_text,
        vad_filter=vad_filter,
    )
    segments = [segment for segment in segments_iter]

    payload = {
        "text": " ".join(segment.text.strip() for segment in segments if segment.text).strip(),
        "language": getattr(info, "language", None),
        "duration": getattr(info, "duration", None),
        "duration_after_vad": getattr(info, "duration_after_vad", None),
        "segments": [
            {
                "start": segment.start,
                "end": segment.end,
                "text": segment.text.strip(),
            }
            for segment in segments
        ],
    }

    print(json.dumps(payload, ensure_ascii=False))


if __name__ == "__main__":
    main()
