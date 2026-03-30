$json = Get-Content -Path 'C:\Users\Bruce\source\repos\ClaudeBudget\gws_json\months\May_payload.json' -Raw -Encoding UTF8
& 'C:\Users\Bruce\AppData\Roaming\npm\gws.cmd' sheets spreadsheets values batchUpdate --params '{"spreadsheetId": "1DEaFJvnXOM_B9GglT6sKXZzq0lCYF__EP_reMm3qu-w"}' --json $json
