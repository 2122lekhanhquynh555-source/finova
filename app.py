import os, io, json, math, secrets, sqlite3, time
from datetime import datetime, timedelta
from functools import wraps
from flask import Flask, render_template, request, jsonify, session, send_file, g
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook

app = Flask(__name__)
app.secret_key = os.getenv('FLASK_SECRET_KEY', 'change-this-secret-before-production-' + secrets.token_hex(8))
DB_PATH = os.getenv('DATABASE_PATH', '/tmp/finova.db')
GEMINI_API_KEY = os.getenv('GOOGLE_AI_STUDIO_API_KEY', '')
GEMINI_MODEL = os.getenv('GEMINI_MODEL', 'gemini-2.5-flash')

SCENARIOS = [
    {'key':'severely_conservative','name':'Severely Conservative','income':-20,'expenses':25,'return':-3,'description':'Severe adverse conditions: a sharp income decline, a major expense surge, and lower assumed returns.'},
    {'key':'conservative','name':'Conservative','income':-10,'expenses':10,'return':0,'description':'Less favorable conditions: a mild income decline, higher expenses, and reduced savings capacity.'},
    {'key':'baseline','name':'Baseline','income':0,'expenses':0,'return':3,'description':'Current user-provided financial baseline with a 3% illustrative annual return assumption.'},
    {'key':'optimistic','name':'Optimistic','income':10,'expenses':0,'return':5,'description':'Improved income and savings capacity with a 5% illustrative annual return assumption.'},
    {'key':'highly_optimistic','name':'Highly Optimistic','income':20,'expenses':-5,'return':7,'description':'Strong income growth and improved expense control with a 7% illustrative annual return assumption.'}
]

FAQS = {
'Getting Started': {
'What is FINOVA?':'FINOVA is a personal finance planning support platform for educational and informational purposes. It combines deterministic calculations with AI explanations to help users explore their financial position and goals.',
'How do I use FINOVA?':'Create an account, complete your profile, enter your current financial information, review your financial assessment, and then use Savings & Investment Planning to compare goals and assumptions.',
'How do I create an account and log in?':'Open the account menu, choose Register, and enter your name, email, and password. Later, use Log in with the same email and password.'},
'Financial Information': {
'How do I enter my income and expenses?':'Enter monthly income and monthly living expenses in VND. Use non-negative values and include recurring obligations consistently.',
'How do I enter my assets and liabilities?':'Enter the current value of your assets and investments, total current debt, and monthly debt repayments. Avoid counting the same asset twice.',
'What should I do if some financial information is missing?':'Do not guess. Complete what you know and identify missing figures before relying on the assessment.'},
'Financial Health Assessment': {
'How does FINOVA assess my financial health?':'FINOVA calculates cash flow, expense-to-income ratio, debt-to-income ratio, savings capacity, assets, liabilities, and estimated net worth from the information you enter.',
'What do cash flow, DTI, and net worth mean?':'Net monthly cash flow is income less living expenses and debt repayments. DTI compares monthly debt repayments with monthly income. Net worth is assets minus current debt.',
'How should I interpret my financial assessment results?':'Treat the results as estimates based on your inputs. Review assumptions, alerts, and scenario differences before making decisions.'},
'Financial Goals': {
'How do I create a financial goal?':'Enter a goal name, target amount, existing funds allocated to it, and a target timeframe. You may also add priority or constraints.',
'How do I set a target amount and timeframe?':'Use the amount you estimate will be required and choose a realistic number of years. These values are planning assumptions, not guarantees.',
'How does FINOVA determine whether my goal is achievable?':'It compares the required monthly savings with your available savings capacity and estimates progress under the assumptions you specify.'},
'Savings & Investment Planning': {
'How much should I save each month?':'FINOVA estimates the monthly amount required to close the goal gap within your target timeframe, then compares it with your available monthly cash flow.',
'How does FINOVA calculate the required monthly savings?':'It calculates the remaining target amount after allocated funds and estimates a monthly contribution over the chosen timeframe. The estimate can include a stated illustrative return assumption.',
'How should I interpret the estimated investment returns?':'Returns are hypothetical assumptions for simulation only. Actual returns can be negative or differ substantially; no return is guaranteed.'},
'Scenarios & What-If Analysis': {
'What are the five financial scenarios?':'Severely Conservative, Conservative, Baseline, Optimistic, and Highly Optimistic. Each uses different assumptions disclosed in the results.',
'How can I change income, expenses, or savings assumptions?':'Use the signed percentage controls and recalculation action in What-If Analysis. The result compares adjusted figures with the current baseline.',
'How do I compare different scenarios?':'Compare cash flow, savings capacity, projected assets, and goal feasibility. The most optimistic scenario is not automatically the most suitable.'},
'Excel Import & Export': {
'How do I upload my financial information using Excel?':'Use a spreadsheet with clearly labeled financial fields. This prototype provides a CSV/JSON-style parsing endpoint for AI-assisted review; verify every imported value before saving.',
'What file format is supported?':'The export function generates .xlsx files. Spreadsheet upload parsing depends on the file and the AI configuration; do not upload sensitive information unless you understand the data handling.',
'How do I export my financial plan and results?':'Use the Export Excel buttons on the assessment or planning pages to download the displayed summary and scenario data.'},
'AI Chatbot': {
"What can FINOVA's AI chatbot help me with?":'It can explain FINOVA features, financial terminology, calculation results, scenario differences, goal planning, and stated risks.',
'How does the chatbot use my financial data?':'When you ask for analysis, relevant profile, goal, and calculated result data may be sent to the configured AI provider. Do not enter information you do not want processed.',
'What should I do if the chatbot gives an unclear answer?':'Ask a focused follow-up question and check the answer against the displayed inputs, formulas, and assumptions. Report unsupported or incorrect answers to support.'},
'Privacy & Security': {
'How does FINOVA protect my financial information?':'This starter application hashes passwords and keeps the AI API key server-side. Before production, configure durable encrypted database storage, HTTPS, backups, access controls, and a privacy policy.',
'Who can access my data?':'In this prototype, account data is stored in the configured application database. The hosting operator and configured service providers may process data as described by their terms. Do not use real sensitive data until production security is reviewed.',
'How can I manage or delete my information?':'Contact the site operator for deletion in this starter version. A production release should provide verified export and deletion controls.'}
}


def db():
    if 'db' not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
    return g.db

@app.teardown_appcontext
def close_db(_e=None):
    conn = g.pop('db', None)
    if conn is not None: conn.close()


def init_db():
    os.makedirs(os.path.dirname(DB_PATH) or '.', exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.executescript('''
    CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY, name TEXT NOT NULL, email TEXT UNIQUE NOT NULL, password_hash TEXT NOT NULL, created_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS profiles(user_id INTEGER PRIMARY KEY, age INTEGER, gender TEXT, country TEXT, city TEXT, phone TEXT, occupation TEXT, updated_at TEXT);
    CREATE TABLE IF NOT EXISTS finances(user_id INTEGER PRIMARY KEY, data TEXT NOT NULL, updated_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS goals(id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL, data TEXT NOT NULL, created_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS reviews(id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL, accuracy INTEGER NOT NULL, groundedness INTEGER NOT NULL, usability INTEGER NOT NULL, processing_time INTEGER NOT NULL, comment TEXT, created_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS tickets(id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, issue_type TEXT, subject TEXT, description TEXT, email TEXT, created_at TEXT NOT NULL);
    ''')
    conn.commit(); conn.close()


def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get('user_id'): return jsonify({'error':'Please log in first.'}), 401
        return fn(*args, **kwargs)
    return wrapper


def current_user():
    if not session.get('user_id'): return None
    return db().execute('SELECT id,name,email FROM users WHERE id=?',(session['user_id'],)).fetchone()


def money(n): return round(float(n or 0), 2)

def compute_financials(d):
    income = money(d.get('monthly_income')); expenses = money(d.get('monthly_expenses')); debtpay = money(d.get('monthly_debt_payments'))
    assets = money(d.get('assets')); investments = money(d.get('investments')); debt = money(d.get('total_debt'))
    cashflow = income - expenses - debtpay
    dti = (debtpay / income * 100) if income > 0 else None
    expense_ratio = (expenses / income * 100) if income > 0 else None
    return {'monthly_income':income,'monthly_expenses':expenses,'monthly_debt_payments':debtpay,'net_monthly_cash_flow':round(cashflow,2),'expense_to_income_ratio':round(expense_ratio,2) if expense_ratio is not None else None,'debt_to_income_ratio':round(dti,2) if dti is not None else None,'monthly_savings_capacity':round(max(0,cashflow),2),'total_assets':round(assets+investments,2),'total_current_liabilities':debt,'estimated_net_worth':round(assets+investments-debt,2)}

def scenario_results(d, years=10):
    base=compute_financials(d); out=[]
    for s in SCENARIOS:
        income=base['monthly_income']*(1+s['income']/100); expenses=base['monthly_expenses']*(1+s['expenses']/100); debtpay=base['monthly_debt_payments']; cf=income-expenses-debtpay
        monthly=max(0,cf); rate=s['return']/100/12; n=max(1,int(years*12)); fv=monthly*n if rate==0 else monthly*((1+rate)**n-1)/rate
        out.append({**s,'adjusted_monthly_income':round(income,2),'adjusted_monthly_expenses':round(expenses,2),'net_monthly_cash_flow':round(cf,2),'monthly_savings_capacity':round(max(0,cf),2),'projected_new_savings':round(fv,2),'projected_net_worth':round(base['estimated_net_worth']+fv,2),'years':years})
    # Suitability score is a transparent heuristic, not an AI-derived probability.
    for r in out:
        score=70 + (10 if r['net_monthly_cash_flow']>=0 else -25) - max(0,(d.get('risk_tolerance')=='low' and r['return']>4)*8) - max(0,abs(r['income'])/5) - max(0,abs(r['expenses'])/6)
        r['suitability_score']=max(20,min(95,round(score)))
        r['suitability_status']='Suitable' if r['suitability_score']>=75 else ('Review assumptions' if r['suitability_score']>=55 else 'Risk contingency')
    return sorted(out,key=lambda x:x['suitability_score'],reverse=True)

def goal_analysis(d, goals):
    f=compute_financials(d); results=[]
    for i,gol in enumerate(goals):
        target=max(0,money(gol.get('target_amount'))); saved=max(0,money(gol.get('allocated_funds'))); years=max(1,int(gol.get('timeframe_years') or 1)); gap=max(0,target-saved); monthly=gap/(years*12); capacity=f['monthly_savings_capacity']; ratio=capacity/monthly if monthly else 1
        score=max(0,min(100,round(ratio*100))) if monthly else 100
        results.append({**gol,'target_amount':target,'allocated_funds':saved,'remaining_amount':gap,'timeframe_years':years,'required_monthly_savings':round(monthly,2),'available_monthly_savings':capacity,'feasibility_score':score,'feasible':capacity>=monthly,'rank':i+1})
    return sorted(results,key=lambda x:(x['feasibility_score'], -x['target_amount']),reverse=True)


def ai_generate(task, payload):
    if not GEMINI_API_KEY:
        return {'text': fallback_ai(task,payload), 'provider':'local-fallback', 'notice':'Google AI Studio key is not configured; deterministic/local guidance shown.'}
    import urllib.request
    prompt = f'''You are FINOVA AI, a careful personal finance education assistant. Task: {task}.\nUse only the supplied data and clearly stated assumptions. Never invent figures, claim calculations you did not perform, guarantee investment returns, or present this as personalized investment advice. If data is missing, say so. Clearly distinguish actual user inputs from hypothetical scenarios. Respond in professional accessible English.\nDATA (JSON): {json.dumps(payload, ensure_ascii=False)}'''
    url=f'https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent?key={GEMINI_API_KEY}'
    body=json.dumps({'contents':[{'parts':[{'text':prompt}]}], 'generationConfig':{'temperature':0.2,'maxOutputTokens':1600}}).encode()
    req=urllib.request.Request(url,data=body,headers={'Content-Type':'application/json'},method='POST')
    try:
        with urllib.request.urlopen(req,timeout=25) as response: data=json.loads(response.read().decode())
        txt=data['candidates'][0]['content']['parts'][0]['text']
        return {'text':txt,'provider':'google-ai-studio'}
    except Exception as e:
        return {'text':fallback_ai(task,payload),'provider':'local-fallback','notice':'AI provider was unavailable; local guidance shown.','detail':str(e)[:180]}

def fallback_ai(task,p):
    f=p.get('financial_summary') or p.get('summary') or {}
    cf=f.get('net_monthly_cash_flow')
    if task=='goal ranking':
        goals=p.get('goals',[])
        if not goals:return 'Add at least one financial goal to receive a comparison. The ranking is an estimate based on the available savings capacity, target amount, and timeframe.'
        first=goals[0]
        return f"Based on the available inputs, **{first.get('name','Primary goal')}** ranks highest under the current simple feasibility comparison, with an estimated feasibility score of {first.get('feasibility_score',0)}%. It requires about {first.get('required_monthly_savings',0):,.0f} VND per month versus estimated available savings of {first.get('available_monthly_savings',0):,.0f} VND. This score is a planning indicator, not a probability of success. Review liquidity needs, debt obligations, and assumptions before choosing."
    if task=='stress test':
        return f"The adjusted net monthly cash flow is {cf:,.0f} VND." if isinstance(cf,(int,float)) else 'Review the adjusted cash flow and savings capacity against the baseline. A negative result means the assumptions create a monthly shortfall; consider lowering expenses, changing the timeline, or revisiting the target. This is an educational estimate.'
    if task=='parse and validate': return 'Input review: verify that all amounts are in VND, monthly and one-time figures are not mixed, assets are not double-counted, and debt repayments are included only once. Missing values should be confirmed rather than guessed. Save the verified values before relying on any projection.'
    return 'I can explain FINOVA features, financial terms, displayed calculations, and scenario assumptions. I only have access to data included in this request. Please share the specific result or ask a focused question. Financial projections are estimates and do not guarantee outcomes.'

@app.route('/')
def index(): return render_template('index.html')

@app.get('/api/me')
def me():
    u=current_user()
    if not u:return jsonify({'user':None})
    p=db().execute('SELECT * FROM profiles WHERE user_id=?',(u['id'],)).fetchone()
    fr=db().execute('SELECT data FROM finances WHERE user_id=?',(u['id'],)).fetchone()
    goals=db().execute('SELECT id,data FROM goals WHERE user_id=? ORDER BY id',(u['id'],)).fetchall()
    return jsonify({'user':dict(u),'profile':dict(p) if p else {},'finance':json.loads(fr['data']) if fr else {},'goals':[{'id':x['id'],**json.loads(x['data'])} for x in goals]})

@app.post('/api/register')
def register():
    x=request.get_json(force=True); name=str(x.get('name','')).strip(); email=str(x.get('email','')).strip().lower(); password=str(x.get('password',''))
    if not name or '@' not in email or len(password)<8 or not any(c.isalpha() for c in password) or not any(c.isdigit() for c in password): return jsonify({'error':'Enter a name, valid email, and a password of at least 8 characters containing letters and numbers.'}),400
    try:
        cur=db().execute('INSERT INTO users(name,email,password_hash,created_at) VALUES(?,?,?,?)',(name,email,generate_password_hash(password),datetime.utcnow().isoformat())); db().commit(); session['user_id']=cur.lastrowid
        db().execute('INSERT OR IGNORE INTO profiles(user_id,updated_at) VALUES(?,?)',(cur.lastrowid,datetime.utcnow().isoformat())); db().commit()
        return jsonify({'ok':True,'user':{'id':cur.lastrowid,'name':name,'email':email}})
    except sqlite3.IntegrityError: return jsonify({'error':'An account with this email already exists.'}),409

@app.post('/api/login')
def login():
    x=request.get_json(force=True); email=str(x.get('email','')).strip().lower(); u=db().execute('SELECT * FROM users WHERE email=?',(email,)).fetchone()
    if not u or not check_password_hash(u['password_hash'],str(x.get('password',''))): return jsonify({'error':'Email or password is incorrect.'}),401
    session.clear(); session['user_id']=u['id']; return jsonify({'ok':True,'user':{'id':u['id'],'name':u['name'],'email':u['email']}})

@app.post('/api/logout')
def logout(): session.clear(); return jsonify({'ok':True})

@app.post('/api/password')
@login_required
def change_password():
    x=request.get_json(force=True); old=str(x.get('current_password','')); new=str(x.get('new_password',''))
    u=db().execute('SELECT password_hash FROM users WHERE id=?',(session['user_id'],)).fetchone()
    if not u or not check_password_hash(u['password_hash'],old): return jsonify({'error':'Current password is incorrect.'}),400
    if len(new)<8 or not any(c.isalpha() for c in new) or not any(c.isdigit() for c in new): return jsonify({'error':'New password must contain at least 8 characters, including letters and numbers.'}),400
    if new!=str(x.get('confirm_password','')): return jsonify({'error':'New passwords do not match.'}),400
    db().execute('UPDATE users SET password_hash=? WHERE id=?',(generate_password_hash(new),session['user_id'])); db().commit()
    return jsonify({'ok':True,'message':'Password updated.'})

@app.post('/api/finance/import')
@login_required
def import_finance():
    f=request.files.get('file')
    if not f or not f.filename.lower().endswith('.xlsx'): return jsonify({'error':'Upload an .xlsx workbook.'}),400
    try:
        wb=__import__('openpyxl').load_workbook(f,read_only=True,data_only=True)
        ws=wb.active; rows=list(ws.iter_rows(values_only=True))
        if len(rows)<2: return jsonify({'error':'The workbook needs a header row and at least one data row.'}),400
        headers=[str(v or '').strip().lower().replace(' ','_') for v in rows[0]]
        values={headers[i]:rows[1][i] for i in range(min(len(headers),len(rows[1])))}
        aliases={'monthly_income':['monthly_income','income','monthly_income_vnd'],'monthly_expenses':['monthly_expenses','expenses','monthly_expenses_vnd'],'investments':['investments','investment_assets'],'assets':['assets','other_assets','total_assets'],'total_debt':['total_debt','liabilities','current_debt'],'monthly_debt_payments':['monthly_debt_payments','debt_repayments','monthly_debt_service']}
        parsed={}
        for field,names in aliases.items():
            raw=next((values[n] for n in names if n in values and values[n] not in (None,'')),0)
            try: parsed[field]=max(0,float(str(raw).replace(',','').replace('VND','').strip()))
            except: parsed[field]=0
        parsed['horizon_years']=10; parsed['liquidity_needs']='medium'; parsed['risk_tolerance']='medium'
        return jsonify({'ok':True,'data':parsed,'review':ai_generate('parse and validate',{'imported_values':parsed,'source_headers':headers})})
    except Exception as e: return jsonify({'error':'Could not read workbook. Check that it is a valid .xlsx file.'}),400

@app.post('/api/reset-request')
def reset_request():
    # Secure default: no account enumeration and no fake reset without an email delivery provider.
    return jsonify({'message':'If the account exists, follow the password-reset instructions configured by the site operator. Email delivery is not configured in this free starter project.'})

@app.post('/api/profile')
@login_required
def save_profile():
    x=request.get_json(force=True); allowed={'age','gender','country','city','phone','occupation'}; vals={k:x.get(k) for k in allowed}
    if vals['age'] not in (None,''):
        try: vals['age']=max(0,min(120,int(vals['age'])))
        except: return jsonify({'error':'Age must be a valid number.'}),400
    db().execute('INSERT INTO profiles(user_id,age,gender,country,city,phone,occupation,updated_at) VALUES(?,?,?,?,?,?,?,?) ON CONFLICT(user_id) DO UPDATE SET age=excluded.age,gender=excluded.gender,country=excluded.country,city=excluded.city,phone=excluded.phone,occupation=excluded.occupation,updated_at=excluded.updated_at',(session['user_id'],vals['age'] or None,vals['gender'],vals['country'],vals['city'],vals['phone'],vals['occupation'],datetime.utcnow().isoformat())); db().commit(); return jsonify({'ok':True})

@app.post('/api/finance')
@login_required
def save_finance():
    x=request.get_json(force=True); fields=['monthly_income','monthly_expenses','investments','assets','total_debt','monthly_debt_payments']
    for k in fields:
        try: x[k]=float(str(x.get(k,0)).replace(',',''))
        except: return jsonify({'error':f'{k} must be a valid amount.'}),400
        if not math.isfinite(x[k]) or x[k]<0: return jsonify({'error':'Money amounts must be zero or positive.'}),400
    x['horizon_years']=max(1,min(30,int(x.get('horizon_years') or 10)))
    x['liquidity_needs']=x.get('liquidity_needs','medium'); x['risk_tolerance']=x.get('risk_tolerance','medium')
    db().execute('INSERT INTO finances(user_id,data,updated_at) VALUES(?,?,?) ON CONFLICT(user_id) DO UPDATE SET data=excluded.data,updated_at=excluded.updated_at',(session['user_id'],json.dumps(x),datetime.utcnow().isoformat())); db().commit()
    return jsonify({'ok':True,'summary':compute_financials(x),'scenarios':scenario_results(x,x['horizon_years'])})

@app.get('/api/assessment')
@login_required
def assessment():
    row=db().execute('SELECT data FROM finances WHERE user_id=?',(session['user_id'],)).fetchone()
    if not row:return jsonify({'error':'Please save your financial information first.'}),400
    d=json.loads(row['data']); f=compute_financials(d); sc=scenario_results(d,d.get('horizon_years',10)); alerts=[]
    if f['net_monthly_cash_flow']<0: alerts.append({'level':'high','title':'Negative cash flow','detail':'Monthly income does not cover living expenses and debt repayments under the current inputs.'})
    if f['debt_to_income_ratio'] is not None and f['debt_to_income_ratio']>=40: alerts.append({'level':'high','title':'High debt repayment burden','detail':'Monthly debt repayments are at least 40% of monthly income. Review affordability and liquidity.'})
    if f['monthly_savings_capacity']<=0: alerts.append({'level':'medium','title':'No current savings capacity','detail':'Current inputs show no positive monthly surplus available for savings.'})
    if money(d.get('assets'))+money(d.get('investments'))<money(d.get('monthly_expenses'))*3: alerts.append({'level':'medium','title':'Potentially limited emergency reserves','detail':'Recorded assets and investments are below three months of living expenses; not all assets may be liquid.'})
    return jsonify({'inputs':d,'summary':f,'scenarios':sc,'alerts':alerts,'assumptions':'Scenario returns and income/expense changes are illustrative assumptions, not forecasts.'})

@app.post('/api/goals')
@login_required
def save_goals():
    x=request.get_json(force=True); goals=x.get('goals',[])
    if not isinstance(goals,list) or len(goals)>20:return jsonify({'error':'Provide a list of up to 20 goals.'}),400
    db().execute('DELETE FROM goals WHERE user_id=?',(session['user_id'],))
    for item in goals:
        name=str(item.get('name','')).strip()
        if not name: continue
        try: target=float(str(item.get('target_amount',0)).replace(',','')); allocated=float(str(item.get('allocated_funds',0)).replace(',','')); years=int(item.get('timeframe_years',1))
        except: return jsonify({'error':'Goal amounts and timeframe must be valid numbers.'}),400
        if target<0 or allocated<0 or years<1 or years>50:return jsonify({'error':'Goal amounts must be non-negative and timeframe must be 1–50 years.'}),400
        payload={'name':name,'target_amount':target,'allocated_funds':allocated,'timeframe_years':years,'priority':item.get('priority','medium'),'notes':item.get('notes','')}
        db().execute('INSERT INTO goals(user_id,data,created_at) VALUES(?,?,?)',(session['user_id'],json.dumps(payload),datetime.utcnow().isoformat()))
    db().commit(); return jsonify({'ok':True})

@app.get('/api/goal-analysis')
@login_required
def get_goal_analysis():
    fr=db().execute('SELECT data FROM finances WHERE user_id=?',(session['user_id'],)).fetchone(); rows=db().execute('SELECT data FROM goals WHERE user_id=?',(session['user_id'],)).fetchall()
    if not fr:return jsonify({'error':'Save financial information first.'}),400
    goals=[json.loads(r['data']) for r in rows]; return jsonify({'goals':goal_analysis(json.loads(fr['data']),goals)})

@app.post('/api/ai/parse')
@login_required
def ai_parse(): return jsonify(ai_generate('parse and validate',request.get_json(force=True)))

@app.post('/api/ai/rank-goals')
@login_required
def ai_rank():
    fr=db().execute('SELECT data FROM finances WHERE user_id=?',(session['user_id'],)).fetchone(); rows=db().execute('SELECT data FROM goals WHERE user_id=?',(session['user_id'],)).fetchall()
    if not fr:return jsonify({'error':'Save financial information first.'}),400
    d=json.loads(fr['data']); goals=goal_analysis(d,[json.loads(r['data']) for r in rows]); return jsonify({'analysis':goals,'ai':ai_generate('goal ranking',{'financial_summary':compute_financials(d),'goals':goals,'risk_tolerance':d.get('risk_tolerance')})})

@app.post('/api/ai/stress')
@login_required
def ai_stress():
    x=request.get_json(force=True); fr=db().execute('SELECT data FROM finances WHERE user_id=?',(session['user_id'],)).fetchone()
    if not fr:return jsonify({'error':'Save financial information first.'}),400
    d=json.loads(fr['data']); baseline=compute_financials(d)
    inc=float(x.get('income_adjustment',0)); exp=float(x.get('expense_adjustment',0)); emergency=float(x.get('unforeseen_expenses',0)); debtadj=float(x.get('debt_adjustment',0)); savingadj=float(x.get('savings_adjustment',0))
    adjusted_income=baseline['monthly_income']*(1+inc/100); adjusted_expenses=baseline['monthly_expenses']*(1+exp/100)+max(0,emergency); adjusted_debt=baseline['monthly_debt_payments']*(1+debtadj/100); cf=adjusted_income-adjusted_expenses-adjusted_debt; capacity=max(0,cf)*(1+savingadj/100)
    result={'baseline':baseline,'adjusted':{'monthly_income':round(adjusted_income,2),'monthly_expenses':round(adjusted_expenses,2),'monthly_debt_payments':round(adjusted_debt,2),'net_monthly_cash_flow':round(cf,2),'monthly_savings_capacity':round(max(0,capacity),2)},'variance':round(cf-baseline['net_monthly_cash_flow'],2),'assumptions':x}
    result['ai']=ai_generate('stress test',{'summary':result['adjusted'],'baseline':baseline,'assumptions':x}); return jsonify(result)

@app.post('/api/ai/chat')
@login_required
def ai_chat():
    x=request.get_json(force=True); question=str(x.get('message','')).strip()
    if not question:return jsonify({'error':'Enter a question.'}),400
    fr=db().execute('SELECT data FROM finances WHERE user_id=?',(session['user_id'],)).fetchone(); rows=db().execute('SELECT data FROM goals WHERE user_id=?',(session['user_id'],)).fetchall()
    payload={'question':question,'financial_summary':compute_financials(json.loads(fr['data'])) if fr else {},'goals':[json.loads(r['data']) for r in rows],'knowledge_base':FAQS,'instructions':'Use only this data and general educational explanations. Do not claim to have accessed unavailable information.'}
    return jsonify(ai_generate('RAG support chatbot',payload))

@app.get('/api/faqs')
def faqs(): return jsonify(FAQS)

@app.post('/api/reviews')
@login_required
def save_review():
    x=request.get_json(force=True)
    vals=[]
    for k in ['accuracy','groundedness','usability','processing_time']:
        try:v=int(x.get(k))
        except:return jsonify({'error':'Please rate all four criteria.'}),400
        if v<1 or v>5:return jsonify({'error':'Ratings must be between 1 and 5 stars.'}),400
        vals.append(v)
    db().execute('INSERT INTO reviews(user_id,accuracy,groundedness,usability,processing_time,comment,created_at) VALUES(?,?,?,?,?,?,?)',(session['user_id'],*vals,str(x.get('comment',''))[:2000],datetime.utcnow().isoformat())); db().commit()
    return jsonify({'ok':True})

@app.get('/api/reviews/summary')
def review_summary():
    r=db().execute('SELECT COUNT(*) n, AVG((accuracy+groundedness+usability+processing_time)/4.0) avg, AVG(accuracy) accuracy, AVG(groundedness) groundedness, AVG(usability) usability, AVG(processing_time) processing_time FROM reviews').fetchone()
    return jsonify(dict(r) if r else {'n':0,'avg':None})

@app.post('/api/tickets')
def ticket():
    x=request.get_json(force=True); required=['issue_type','subject','description','email']
    if any(not str(x.get(k,'')).strip() for k in required):return jsonify({'error':'Complete issue type, subject, description, and email.'}),400
    db().execute('INSERT INTO tickets(user_id,issue_type,subject,description,email,created_at) VALUES(?,?,?,?,?,?)',(session.get('user_id'),str(x['issue_type'])[:120],str(x['subject'])[:200],str(x['description'])[:5000],str(x['email'])[:250],datetime.utcnow().isoformat())); db().commit(); return jsonify({'ok':True,'message':'Your request has been recorded in this prototype.'})

@app.post('/api/export')
@login_required
def export_xlsx():
    x=request.get_json(force=True); wb=Workbook(); ws=wb.active; ws.title='Financial Summary'
    summary=x.get('summary',{}); ws.append(['Metric','Value','Unit'])
    labels={'monthly_income':'Total Monthly Income','monthly_expenses':'Total Monthly Expenses','net_monthly_cash_flow':'Net Monthly Cash Flow','expense_to_income_ratio':'Expense-to-Income Ratio','debt_to_income_ratio':'Debt-to-Income Ratio','monthly_savings_capacity':'Monthly Savings Capacity','total_assets':'Total Assets','total_current_liabilities':'Total Current Liabilities','estimated_net_worth':'Estimated Net Worth'}
    for k,label in labels.items(): ws.append([label,summary.get(k,''),'%' if 'ratio' in k else 'VND'])
    sc=wb.create_sheet('Scenarios'); sc.append(['Scenario','Income change %','Expense change %','Assumed annual return %','Net monthly cash flow (VND)','Monthly savings capacity (VND)','Projected new savings (VND)','Suitability score (heuristic)'])
    for s in x.get('scenarios',[]):sc.append([s.get('name'),s.get('income'),s.get('expenses'),s.get('return'),s.get('net_monthly_cash_flow'),s.get('monthly_savings_capacity'),s.get('projected_new_savings'),s.get('suitability_score')])
    goals=wb.create_sheet('Goals'); goals.append(['Goal','Target amount (VND)','Allocated funds (VND)','Years','Required monthly savings (VND)','Available monthly savings (VND)','Feasibility score (not probability)'])
    for a in x.get('goals',[]): goals.append([a.get('name'),a.get('target_amount'),a.get('allocated_funds'),a.get('timeframe_years'),a.get('required_monthly_savings'),a.get('available_monthly_savings'),a.get('feasibility_score')])
    for sheet in wb.worksheets:
        sheet.freeze_panes='A2'; sheet.auto_filter.ref=sheet.dimensions
        for col in sheet.columns:
            width=min(42,max(14,max(len(str(c.value or '')) for c in col)+2)); sheet.column_dimensions[col[0].column_letter].width=width
    bio=io.BytesIO(); wb.save(bio); bio.seek(0)
    return send_file(bio,as_attachment=True,download_name='FINOVA_Financial_Report.xlsx',mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

with app.app_context(): init_db()

if __name__=='__main__': app.run(host='0.0.0.0',port=int(os.getenv('PORT','5000')),debug=False)
