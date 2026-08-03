# EDC Form Filler (web version)

Same logic as the original `main.py` script, wrapped in a small Flask app.
Instead of running from the terminal and reading `t.txt` off disk, you paste
the text into a web page. Anything the script can't find automatically, it
asks you for on the next page — same fields it used to prompt for in the
terminal.

## Setup

1. Put this folder's `app.py`, the `templates/` folder, and your
   **`EDC Form.docx`** template all in the same directory. The app looks for
   `EDC Form.docx` right next to `app.py`, exactly like the original script
   looked for it in the current working directory.

2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

3. Run it:

   ```bash
   python app.py
   ```

4. Open **http://127.0.0.1:5000** in your browser.

## How it works

1. **Paste page** – paste the full contents of your `t.txt` file (multiple
   records separated by a line that just says `another` still work exactly
   as before).
2. **Review page** – for each record you'll see:
   - The fields it found automatically (name, mobile, village, union,
     upazila, router serial/MAC, etc.)
   - Any fields it *couldn't* find, as required inputs (highlighted) — this
     mirrors the old `prompt_for_missing_fields()` step.
   - The fields the original script always asked for manually every time
     (EDC Book No., NMS ID, institution code, cable brand/qty, lat-long).
3. Click **Generate** — it fills `EDC Form.docx` for each record the same
   way the original script did (fonts, header line, cable qty row, etc. all
   unchanged) and gives you a `.docx` download (or a `.zip` if there were
   multiple records).

## Notes

- `app.secret_key` in `app.py` is just used to sign the session cookie —
  change it to any random string before you use this beyond your own
  machine.
- Parsed records are held in memory between the two pages (keyed by a token
  in your session cookie), so it's fine for one person using it locally.
  If you ever host this for multiple simultaneous users, swap the in-memory
  `STORE` dict for a real store (e.g. Redis).
