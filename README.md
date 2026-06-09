# FCI Deposit Converter — Web

A small Vercel project: drop an FCI Notification-of-Deposit PDF in the browser,
get back a formatted Excel workbook.

## Files

```
.
├── index.html              Static frontend, served at /
├── api/
│   ├── convert.py          Serverless function at /api/convert
│   └── deposit_to_excel.py PDF → XLSX logic (same code as the desktop app)
├── requirements.txt        Python deps for the serverless function
├── vercel.json             Function config (max duration)
└── README.md
```

The frontend POSTs the raw PDF bytes (no multipart) to `/api/convert` with an
`X-Filename` header. The function writes the bytes to a temp directory, calls
`convert()`, reads the resulting .xlsx, and returns it with an
`attachment; filename=…` Content-Disposition header. Temp files are deleted
before the function returns.

## Deploy

You'll need a free Vercel account: <https://vercel.com/signup>.

### Option 1 — Vercel CLI (fastest)

```bash
npm i -g vercel        # one-time
cd fci-web
vercel                 # follow prompts; choose "link to new project"
vercel --prod          # deploy production
```

Vercel will detect the Python function in `api/` automatically and install
the dependencies from `requirements.txt`. The first deploy takes ~60 seconds
because it has to provision the Python runtime; subsequent deploys are fast.

You'll get a URL like `https://fci-deposit-converter.vercel.app`. Done.

### Option 2 — GitHub integration

1. Push this folder to a new GitHub repo (public or private both work).
2. In the Vercel dashboard, click "Add New… → Project", import the repo.
3. Accept the defaults (Framework Preset: "Other"). Click Deploy.

Every push to `main` will redeploy automatically.

## Local development

```bash
npm i -g vercel
cd fci-web
vercel dev
```

That spins up the static site and the Python function on `http://localhost:3000`,
mimicking the production setup. Useful for testing changes before redeploying.

## Limits (free tier)

| Limit                  | Value     |
|------------------------|-----------|
| Request body           | 4.5 MB    |
| Function execution     | 10 s      |
| Function memory        | 1024 MB   |
| Bandwidth              | 100 GB/mo |

The 4.5 MB request body limit is the practical ceiling on PDF size. The
sample PDF (25 pages, 159 KB) parses in ~3 seconds, so the 10s limit is
fine for normal deposit notifications. The Pro plan ($20/mo) bumps both
ceilings if you ever need bigger PDFs or longer execution time.

## Notes

- The frontend's "files are not stored or logged" claim is true for the
  function as written: it uses a per-request `tempfile.mkdtemp()` and
  `shutil.rmtree()` in a `finally`, and no logging beyond Python's
  `print()` (which Vercel captures to its own logs without the file
  contents). If you want to harden this further, set the Vercel project
  to not persist function logs.
- The Python function loads pdfplumber + openpyxl on cold start. Cold
  starts take ~1.5–2 seconds; warm calls are ~3 seconds (the parse
  itself). To keep things consistently fast, you'd need to use Vercel's
  paid "Fluid Compute" feature; the free tier is fine for occasional use.
- The PDF parser is tuned to FCI's specific template (column X-positions
  in the Centurion-generated layout). If FCI ever changes the template,
  edit the `COLUMNS` list at the top of `deposit_to_excel.py`.
