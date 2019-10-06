# srxPolicyExport
This script will take in 'show security | dis inh | dis set' and 'show security policy hits' output and make pretty XLS sheets from the security rules. It will resolve address book and address book sets and output both.

Output column names:

```
policy-name
policy-desc
policy-hits
from-zone
to-zone
source-address
source-address-resolved
destination-address
destination-address-resolved
application
then
```


Usage
-----
create two files, one with all the security sets and one with the policy hit counters. them define them at the top of the main script.


The output will be a xls with the first sheet as all the rules and the following sheets the break of of zone to zone.



Notes
-----
- Hit count only works if the policy is already set to count, obviously.
- Rules will be in the other they are outputed so should stay in the processed order.


Limitations
-----------
- Global Policy is not yet supported.
- Only policy statements desc|match|then are supported (so no schedule).
- Applicaions are not yet resolved.
- sheet names have a 31 character limit, anything longer will be truncated by exceljs automagically. fromZoneName->ToZoneName is used for sheet names.
