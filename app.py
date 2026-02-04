from flask import Flask, request, send_file, render_template_string
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from io import BytesIO

app = Flask(__name__)

HTML = open('/tmp/html.txt').read() if False else '''<!DOCTYPE html>
<html>
<head>
<title>Visa Excel Generator</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:linear-gradient(135deg,#667eea,#764ba2);min-height:100vh;padding:20px}
.container{max-width:1000px;margin:0 auto;background:#fff;border-radius:20px;padding:40px;box-shadow:0 20px 60px rgba(0,0,0,0.3)}
h1{color:#333;margin-bottom:10px;font-size:2rem}
h2{color:#333;margin:20px 0 15px 0;font-size:1.5rem}
.form-row{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:20px}
label{display:block;margin-bottom:5px;color:#555;font-weight:600;font-size:0.9rem}
input,select{width:100%;padding:12px;border:2px solid #e0e0e0;border-radius:8px;font-size:1rem;transition:border 0.3s}
input:focus,select:focus{outline:none;border-color:#667eea}
button{background:linear-gradient(135deg,#667eea,#764ba2);color:#fff;border:none;padding:14px 28px;border-radius:8px;cursor:pointer;margin-right:10px;font-size:1rem;font-weight:600;transition:transform 0.2s}
button:hover{transform:translateY(-2px)}
.card{background:#f8f9fa;border-radius:12px;padding:20px;margin:10px 0;border-left:4px solid #667eea;display:flex;justify-content:space-between;align-items:center}
.remove{background:#ff4757;padding:8px 16px;border-radius:6px}
.remove:hover{background:#ff3838}
@media(max-width:768px){.form-row{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="container">
<h1>🛂 Visa Excel Generator</h1>
<div class="form-row">
<div><label>First Name *</label><input id="firstName" required></div>
<div><label>Last Name *</label><input id="lastName" required></div>
</div>
<div class="form-row">
<div><label>Category</label><select id="category"><option>National Visa</option><option>Schengen Visa</option></select></div>
<div><label>Subcategory</label><select id="subcategory"><option>Work Visa</option></select></div>
</div>
<div class="form-row">
<div><label>Passport Number</label><input id="passport"></div>
<div><label>Birthdate (dd.mm.yyyy)</label><input id="birthdate" placeholder="13.03.1994"></div>
</div>
<div class="form-row">
<div><label>Passport Validity (dd.mm.yyyy)</label><input id="passportValidity" placeholder="12.11.2029"></div>
<div><label>Gender</label><select id="gender"><option>Male</option><option>Female</option></select></div>
</div>
<div class="form-row">
<div><label>Phone (with country code)</label><input id="phone" placeholder="245857456140"></div>
<div><label>Nationality</label><select id="nationality"><option>GUINEA-BISSAU</option><option>SENEGAL</option></select></div>
</div>
<div class="form-row">
<div><label>Book Date From (dd.mm.yyyy)</label><input id="bookDateFrom" placeholder="01.01.2025"></div>
<div><label>Book Date To (dd.mm.yyyy)</label><input id="bookDateTo" placeholder="31.01.2025"></div>
</div>
<div class="form-row">
<div><label>Agent Name *</label><input id="agentName" required></div>
<div><label>Days Gap</label><input id="daysGap" type="number"></div>
</div>
<div class="form-row">
<div><label>Price</label><input id="price" type="number"></div>
<div><label>Group</label><input id="group"></div>
</div>
<div class="form-row">
<div><label>Email</label><input id="email" type="email"></div>
<div></div>
</div>
<button onclick="add()">➕ Add</button>
<div id="list" style="display:none;margin-top:30px">
<h2>Customers (<span id="cnt">0</span>)</h2>
<div id="cust"></div>
<button onclick="gen()">📥 Download</button>
</div>
</div>
<script>
let c=[];
const s={'National Visa':['Work Visa','Job Seeker Visa','Medical Treatment Visa','Study Visa','Family Reunion Visa'],'Schengen Visa':['Schengen Visa']};
document.getElementById('category').onchange=function(){document.getElementById('subcategory').innerHTML=s[this.value].map(v=>`<option>${v}</option>`).join('')};
function add(){
const f=document.getElementById('firstName').value;
const l=document.getElementById('lastName').value;
const a=document.getElementById('agentName').value;
if(!f||!l||!a){alert('Fill required fields');return}
c.push({
firstName:f,
lastName:l,
category:document.getElementById('category').value,
subcategory:document.getElementById('subcategory').value,
agentName:a,
passport:document.getElementById('passport').value,
birthdate:document.getElementById('birthdate').value,
passportValidity:document.getElementById('passportValidity').value,
gender:document.getElementById('gender').value,
phone:document.getElementById('phone').value,
nationality:document.getElementById('nationality').value,
bookDateFrom:document.getElementById('bookDateFrom').value,
bookDateTo:document.getElementById('bookDateTo').value,
daysGap:document.getElementById('daysGap').value,
price:document.getElementById('price').value,
group:document.getElementById('group').value,
email:document.getElementById('email').value
});
document.getElementById('cnt').textContent=c.length;
document.getElementById('list').style.display='block';
document.getElementById('cust').innerHTML=c.map((x,i)=>`<div class="card"><div><b>${x.firstName} ${x.lastName}</b><br>${x.category} • ${x.subcategory}</div><button class="remove" onclick="c.splice(${i},1);add()">Remove</button></div>`).join('');
document.getElementById('firstName').value='';
document.getElementById('lastName').value='';
document.getElementById('agentName').value='';
document.getElementById('passport').value='';
document.getElementById('birthdate').value='';
document.getElementById('passportValidity').value='';
document.getElementById('phone').value='';
document.getElementById('bookDateFrom').value='';
document.getElementById('bookDateTo').value='';
document.getElementById('daysGap').value='';
document.getElementById('price').value='';
document.getElementById('group').value='';
document.getElementById('email').value='';
}
async function gen(){
const r=await fetch('/generate',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({customers:c})});
if(r.ok){
const b=await r.blob();
const u=URL.createObjectURL(b);
const a=document.createElement('a');
a.href=u;
a.download='GUINEA_VISA.xlsx';
a.click();
alert('✅ Downloaded!');
}else{
alert('❌ Error');
}
}
</script>
</body>
</html>'''

@app.route('/')
def index():
    return render_template_string(HTML)

@app.route('/generate', methods=['POST'])
def generate():
    try:
        customers = request.json.get('customers', [])
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        
        # Create sheets
        ds = wb.create_sheet('Data', 0)
        s2 = wb.create_sheet('Лист2', 1)
        s2.sheet_state = 'hidden'
        s3 = wb.create_sheet('Лист3', 2)
        s3.sheet_state = 'hidden'
        
        # Setup Лист2
        s2['A1']='City';s2['B1']='Category';s2['C1']='National Visa Subcategory';s2['D1']='Schengen Visa Subcategory'
        s2['A2']='Bissau';s2['B2']='National Visa';s2['C2']='Work Visa';s2['D2']='Schengen Visa'
        s2['B3']='Schengen Visa';s2['C3']='Job Seeker Visa'
        s2['C4']='Medical Treatment Visa';s2['C5']='Study Visa';s2['C6']='Family Reunion Visa'
        
        # Setup Лист3
        s3['A1']='Gender';s3['B1']='Country'
        s3['A2']='Male';s3['B2']='GUINEA-BISSAU'
        s3['A3']='Female';s3['B3']='SENEGAL'
        
        # Named ranges
        wb.defined_names['City']=DefinedName('City',attr_text="'Лист2'!$A$2")
        wb.defined_names['Category']=DefinedName('Category',attr_text="'Лист2'!$B$2:$B$3")
        wb.defined_names['National_Visa_Subcategory']=DefinedName('National_Visa_Subcategory',attr_text="'Лист2'!$C$2:$C$6")
        wb.defined_names['Schengen_Visa_Subcategory']=DefinedName('Schengen_Visa_Subcategory',attr_text="'Лист2'!$D$2")
        wb.defined_names['Gender']=DefinedName('Gender',attr_text="'Лист3'!$A$2:$A$3")
        wb.defined_names['Country']=DefinedName('Country',attr_text="'Лист3'!$B$2:$B$3")
        
        # Headers
        ds.append(['City','Category','Subcategory','Price','Last Name','First Name','Passport number','Birthdate (dd.mm.yyyy)','Passport validity  (dd.mm.yyyy)','Gender (M/F)','Phone (with country code)','Nationality','Book date from  (dd.mm.yyyy)','Book date to  (dd.mm.yyyy)','Agent Name (required)','Days gap','Group (not required)','e-mail'])
        
        # Column widths
        ws=[9.57,16.14,18.71,5.14,15.71,25.43,15.43,12.14,13.14,6.57,23.86,9.43,12.43,11.86,16.5,5.43,10.43,24.43]
        for i,w in enumerate(ws,1):
            ds.column_dimensions[openpyxl.utils.get_column_letter(i)].width=w
        
        ds.row_dimensions[1].height=33.0
        
        # Add customers
        for c in customers:
            ds.append([
                'Bissau',
                c.get('category','National Visa'),
                c.get('subcategory','Work Visa'),
                c.get('price',''),
                c.get('lastName',''),
                c.get('firstName',''),
                c.get('passport',''),
                c.get('birthdate',''),
                c.get('passportValidity',''),
                c.get('gender','Male'),
                c.get('phone',''),
                c.get('nationality','GUINEA-BISSAU'),
                c.get('bookDateFrom',''),
                c.get('bookDateTo',''),
                c.get('agentName',''),
                c.get('daysGap',''),
                c.get('group',''),
                c.get('email','')
            ])
        
        # Data validations
        dv1=DataValidation(type="list",formula1='City',allow_blank=True)
        ds.add_data_validation(dv1)
        dv1.add('A2:A202')
        
        dv2=DataValidation(type="list",formula1='Category',allow_blank=True)
        ds.add_data_validation(dv2)
        dv2.add('B2:B202')
        
        dv3=DataValidation(type="list",formula1='INDIRECT(SUBSTITUTE(SUBSTITUTE(B2," ","_"),"-","_")&"_Subcategory")',allow_blank=True)
        ds.add_data_validation(dv3)
        dv3.add('C2:C202')
        
        dv4=DataValidation(type="list",formula1='Gender',allow_blank=True)
        ds.add_data_validation(dv4)
        dv4.add('J2:J202')
        
        dv5=DataValidation(type="list",formula1='Country',allow_blank=True)
        ds.add_data_validation(dv5)
        dv5.add('L2:L202')
        
        # Save
        out=BytesIO()
        wb.save(out)
        out.seek(0)
        
        return send_file(out,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',as_attachment=True,download_name='GUINEA_VISA.xlsx')
    except Exception as e:
        return {'error':str(e)},500

if __name__=='__main__':
    app.run(debug=True,host='0.0.0.0',port=5000)
