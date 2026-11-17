# Kites certificate generator

Create the Kites Robo Challenge certificate of participation and keep a participant list.

The certificate reads:

Choose Participation or Winner. Only the name, institution and grade change. A winner certificate says the person is a winner, and its heading reads Certificate of Winner. The phone number is stored in the participant list only.

## Run it on this computer

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python app.py
```

Open http://127.0.0.1:8765.

The first certificate creates `Kites-Robo-Challenge.xlsx` in this folder. Later certificates add a row to that same file.
