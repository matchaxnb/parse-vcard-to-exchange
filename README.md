# parse-vcard

This tool parses a collection of vcards, such as those extracted from Roundcube,
and generates a CSV that is fit to be imported on Outlook.

# Tips

## Export vcards from roundcube

```mysql
USE roundcube
tee contacts.lst
SELECT c.vcard contact_vcard
FROM 
    contacts c INNER JOIN
    contactgroupmembers cgm ON cgm.contact_id = c.contact_id INNER JOIN
    contactgroups cg ON cgm.contactgroup_id = cg.contactgroup_id
WHERE
    c.del <> 1
GROUP BY c.contact_id\G
```

```bash
grep -v 'row ***' contacts.lst > /tmp/f
mv /tmp/f contacts.lst
```

## Convert CSV

```bash
cat contacts.lst | python parse_vcards.py contacts.csv
```

## Import to Office 365

Follow [this guide](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell?view=exchange-ps) to log-in in a PowerShell session.

Then follow [this guide](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell?view=exchange-ps) to import the CSV.

If you need to purge all contacts, after logging-in, do this in PowerShell:

```powershell
Get-MailContact | Remove-MailContact
```

## Fill placeholder e-mails when this info is missing (pity)

```bash
N=1
cat contacts.lst | while read l; do N=$(($N + 1)); echo $l | sed 's!^,!'$N'@missing.email.invalid,!g' ; done  > /tmp/f;
mv /tmp/f contacts.lst
```

With this you should be all set.

This is quite minimal, but it does the job for bulk import.
