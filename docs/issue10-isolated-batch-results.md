# Issue 10 isolated batch acceptance

Tested on 2026-11-19 using the independent Render service
`duty-issue10-isolated`, Free plan in Oregon, one Gunicorn sync worker and
`--timeout 120`. The service receives approved fixture copies in authenticated
HTTP requests and has no production Supabase configuration.

Each successful response ran the existing per-document delivery validation,
returned a ZIP with the exact requested dates, passed ZIP CRC checks, and opened
every DOCX. Every internal DOCX part matched the pre-optimization baseline.

| Days | Run 1 seconds | Run 2 seconds | Run 3 seconds | Success |
| --- | ---: | ---: | ---: | --- |
| 5 | 29.993 | 37.275 | 23.710 | 3/3 |
| 7 | 39.779 | 51.148 | 31.599 | 3/3 |
| 10 | 50.206 | 58.688 | 41.262 | 3/3 |

Run 2 in each group adds two seconds to every simulated Storage read.
The three ten-day groups cover November 1–30. The five- and seven-day groups
exercise template transitions and the previously failing dates.

After adding the ten-day limit, an eleven-day HTTP request returned 400 with
no attachment. A final ten-day mixed-template run covering November
1, 2, 7, 10, 13, 14, 20, 21, 22 and 28, also with two seconds per read,
returned HTTP 200 in 61.233 seconds. Its worker peak RSS was 196412 KiB
(191.8 MiB). The earlier sequential runs reported a lifetime worker high-water
mark up to 471016 KiB (460.0 MiB), not an independent per-request peak.

The selected limit is ten days, the largest tested size. The slowest observed
successful run leaves 58.767 seconds below the 120-second timeout. This is an
observed margin, not a guarantee under all Free-plan contention or real Storage
latency. Production Storage latency was not exercised.

Local checks: Python 52/52, frontend 15/15, and git diff whitespace checks passed.
PR 11 remains a draft. No production application deployment is part of this test.
