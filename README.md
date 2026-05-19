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

For full documentation on modes, architecture, configuration, and API reference, see [info.md](info.md).
