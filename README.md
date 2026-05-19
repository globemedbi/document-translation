# PDF Form Translator

## Quickstart

1. Open `docker-compose.yml` and replace `<your_api_key>` with your OpenAI API key
2. Run:

```bash
docker compose up --build
```

| Service | URL |
|---|---|
| Streamlit UI | http://localhost:8501 |
| FastAPI backend | http://localhost:8000 |
| API docs | http://localhost:8000/docs |

## Models used

| Env var | Default | Role |
|---|---|---|
| `OPENAI_MODEL` | `gpt-5.5` | Extraction, translation, observer |
| `OPENAI_HEAVY_MODEL` | `o4-mini` | Layout planner |
| `OPENAI_CODING_MODEL` | `o3` | Code generation |

For full documentation on modes, architecture, configuration, and API reference, see [info.md](info.md).
