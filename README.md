# Quiz Spark

A Streamlit app that creates cheerful, grade-appropriate quizzes with OpenRouter and
exports them as Microsoft Word (`.docx`) worksheets with an answer key.

## Set up

The project virtual environment is `.venv`. From this folder, activate it in
PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

Install dependencies:

```powershell
python -m pip install -r requirements.txt
```

Copy `.env.example` to `.env`, then set `OPENROUTER_API_KEY` to your API key.
The default model, `openrouter/free`, routes requests to an available free model.
You can instead set `OPENROUTER_MODEL` to a specific model ID available to your
OpenRouter account.

- `OPENROUTER_BASE_URL`: defaults to `https://openrouter.ai/api/v1`.
- `OPENROUTER_MODEL`: defaults to `openrouter/free`.
- `OPENROUTER_API_KEY`: your OpenRouter API key. The app temporarily falls back
  to `OLLAMA_API_KEY` if this variable is not set, so existing `.env` files keep
  working; prefer renaming it to `OPENROUTER_API_KEY`.

Keep `.env` private; it is excluded from Git.

## Run

```powershell
streamlit run app.py
```

Choose a student grade, subject, topic, challenge level, question style, and
question count. Generate a preview and download a Word file containing the
student worksheet and a separate answer key.
