# domain-report-sheet-python

This Python script reads domain names from an Excel file and creates a output excel file with table of domains with their registrar, domain expiry date, and important DNS records details. The script uses the `openpyxl` library to read data from the Excel file, and the `whois` library to fetch domain details.

| Domain Name|	HTTP status	|Registrar	|Expiry Date|	A record	|MX records|	TXT record|	DKIM record
| ------ | ------ |------ | ------ |------ | ------ |------ | ------ |
| google.com | | | | | | | |
| yahoo.com | | | | | | | |
| facebook.com | | | | | | | |

#

## Dependencies

- Python 3.6+
- `openpyxl` library (`pip install openpyxl`)
- `whois` library (`pip install python-whois`)


```

## Notes

- The script only fetches details for domains with valid TLDs.
- The `whois` library has rate limiting and may throw errors if called too frequently. Consider using a different library or adding a delay between requests if you need to fetch details for a large number of domains.
