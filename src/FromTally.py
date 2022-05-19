import json

import requests
import xmltodict
import pandas as pd
from pprintpp import pprint as pp


data = """<ENVELOPE>
<HEADER>
<VERSION>1</VERSION>
<TALLYREQUEST>Export</TALLYREQUEST>
<TYPE>Data</TYPE>
<ID>Ledger Vouchers</ID>
</HEADER>
<BODY>
<DESC>
<STATICVARIABLES>

<SVEXPORTFORMAT>$$SysName:ASCII</SVEXPORTFORMAT>
<ledgername>onlineSale</ledgername>

<!--Specify the Voucher Type here-->
<VOUCHERTYPENAME>Sales</VOUCHERTYPENAME>

<!--Specify the Export format here  HTML or XML or SDF-->
<SVEXPORTFORMAT>$$SysName:HTML</SVEXPORTFORMAT>

<!--Set the Columnar format variable here -->
<COLUMNARDAYBOOK>Yes</COLUMNARDAYBOOK>

<!--Set the SVColumntype variable here -->
<SVCOLUMNTYPE>$$SysName:AllItems</SVCOLUMNTYPE>


</STATICVARIABLES>
</DESC>
</BODY>
</ENVELOPE>"""


url = 'http://localhost:9000/'
r = requests.post(url=url, data=data)

if r.status_code == 200:
    text = r.text.split("\r\n ")
    pp(text)
else:
    pp(r.status_code)

