# NanaBot, Zagreb SSS job alerts- "ScraperBot za državne i gradske poslove, mjesto rada :Zagreb, stručna sprema :SSS. Samo ponuda javnih poslova po kriterijima!!!"
Nažalost , bot nikad ne ide van u određeno vrijeme zbog gužve valjda na serverima, ali dnevne obavijesti stižu, nekad jedna, nekad 3 :)....koga briga, radi...

Automated Job Scraper with Email Notifications.
Daily check for high-school (`SSS`) public jobs in Zagreb. The script reads three city boards and the central state portal, emails new matches through [Resend](https://resend.com), and records what it already sent in `sent_jobs.json`.

## Sources

- City offices: https://zagreb.hr/objavljeni-natjecaji-oglasi/6501
- Culture institutions: https://zagreb.hr/javni-natjecaji-za-zaposljavanje-u-ustanovama-u-ku/202452
- Kindergartens: https://vrtici.zagreb.hr/natjecaji-za-zaposljavanje/122
- State portal: https://selekcija.gov.hr/natjecaji/objavljeni-natjecaji

A job is kept when the notice asks for `SSS`, `srednja stručna sprema`, `završena srednja škola`, or HKO level 4.2, and the workplace is Zagreb. The first successful run emails every current match. Later runs email only ids that are not yet in `sent_jobs.json`.

## GitHub Actions

Add these repository secrets:

- `RESEND_API_KEY` — Resend API key
- `RECEIVER_EMAIL` — address that should receive the alert
- `RESEND_FROM_EMAIL` — verified sender, for example `NanaBot <jobs@yourdomain.com>`

The workflow [`.github/workflows/daily-jobs.yml`](.github/workflows/daily-jobs.yml) runs at 06:00 and 07:00 UTC. One of those is 08:00 in Zagreb (the offset changes with daylight saving time). If the morning run already saved the new ids, the second run does not send another email.

After a successful send, the workflow commits `sent_jobs.json` as `github-actions[bot]`. A failed send leaves the file unchanged, so the same jobs are tried again next time.

You can start a run manually from the Actions tab with **Run workflow**.

## Local run

```powershell
python -m pip install -r requirements.txt
$env:RESEND_API_KEY = "re_..."
$env:RECEIVER_EMAIL = "sister@example.com"
$env:RESEND_FROM_EMAIL = "NanaBot <jobs@yourdomain.com>"
python scrape_jobs.py
```

Copy [`.env.example`](.env.example) only as a reminder. The script reads `os.environ`. It does not load a `.env` file.
