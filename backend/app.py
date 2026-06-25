from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import pandas as pd
import numpy as np
import json, io, os, warnings
warnings.filterwarnings('ignore')

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                 TableStyle, HRFlowable, PageBreak, KeepTogether)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from io import BytesIO
from reportlab.platypus import Image as RLImage

app = Flask(__name__)
CORS(app)

# ──────────────────────────────────────────────
#  FILE LOADER
# ──────────────────────────────────────────────
def load_file(file) -> pd.DataFrame:
    name = file.filename.lower()
    content = file.read()
    if name.endswith('.csv'):
        df = pd.read_csv(io.BytesIO(content), encoding='utf-8', on_bad_lines='skip')
    elif name.endswith('.xlsx'):
        df = pd.read_excel(io.BytesIO(content), engine='openpyxl')
    elif name.endswith('.xls'):
        df = pd.read_excel(io.BytesIO(content))
    elif name.endswith('.json'):
        data = json.loads(content)
        df = pd.DataFrame(data) if isinstance(data, list) else pd.json_normalize(data)
    else:
        raise ValueError(f"Unsupported file: {name}")
    df.columns = [str(c).strip().lower().replace(' ','_') for c in df.columns]
    return df

# ──────────────────────────────────────────────
#  COLUMN DETECTION
# ──────────────────────────────────────────────
def detect_cols(df):
    cols = df.columns.tolist()
    def find(*keys):
        for k in keys:
            for c in cols:
                if k in c: return c
        return None
    detected = {
        'customer': find('customer_id','cust_id','client_id','user_id','customer'),
        'name':     find('customer_name','name','client_name','full_name'),
        'sales':    find('sales','revenue','amount','total','price','sale_amount'),
        'product':  find('product','item','product_name','item_name','sku'),
        'category': find('category','cat','product_category'),
        'region':   find('region','location','area','zone','city','state','country'),
        'date':     find('date','order_date','purchase_date','created_at','timestamp'),
        'quantity': find('quantity','qty','units','count'),
        'ctype':    find('customer_type','cust_type'),
        'profit':   find('profit','margin','net'),
        'order_id': find('order_id','order','transaction_id'),
    }
    
    # Smart Fallbacks
    num_cols = []
    str_cols = []
    for c in cols:
        try:
            converted = pd.to_numeric(df[c].dropna().head(10).astype(str).str.replace(r'[^\d.]','',regex=True), errors='coerce')
            if converted.notnull().mean() > 0.5:
                num_cols.append(c)
            else:
                str_cols.append(c)
        except:
            str_cols.append(c)
            
    if not detected['sales'] and num_cols:
        detected['sales'] = num_cols[0]
    if not detected['quantity'] and len(num_cols) > 1:
        if num_cols[0] == detected['sales']:
            detected['quantity'] = num_cols[1]
        else:
            detected['quantity'] = num_cols[0]
    if not detected['customer'] and not detected['name'] and str_cols:
        detected['customer'] = str_cols[0]
    if not detected['product'] and str_cols:
        candidates = [c for c in str_cols if c not in (detected['customer'], detected['name'])]
        detected['product'] = candidates[0] if candidates else str_cols[0]
    if not detected['date']:
        for c in cols:
            if any(k in c for k in ('year', 'month', 'day', 'time')):
                detected['date'] = c
                break
        else:
            for c in cols:
                try:
                    converted = pd.to_datetime(df[c].dropna().head(5), errors='coerce')
                    if converted.notnull().mean() > 0.6:
                        detected['date'] = c
                        break
                except:
                    pass
    if not detected['sales'] and cols:
        detected['sales'] = cols[0]
    if not detected['customer'] and len(cols) > 1:
        detected['customer'] = cols[1]
    elif not detected['customer'] and cols:
        detected['customer'] = cols[0]
        
    return detected

def fmtm(n):
    try:
        n = float(n)
        if n >= 1_000_000: return f"${n/1_000_000:.2f}M"
        if n >= 1_000:     return f"${n/1_000:.1f}K"
        return f"${n:,.2f}"
    except: return "-"

# ──────────────────────────────────────────────
#  CORE ANALYTICS
# ──────────────────────────────────────────────
def process_dataframe(df, cols):
    sales_col = cols['sales']
    qty_col   = cols['quantity']
    cust_col  = cols['customer'] or cols['name']

    if sales_col:
        df[sales_col] = pd.to_numeric(
            df[sales_col].astype(str).str.replace(r'[^\d.]','',regex=True), errors='coerce').fillna(0)
    if qty_col:
        df[qty_col] = pd.to_numeric(df[qty_col], errors='coerce').fillna(1)

    customers = {}
    if cust_col and sales_col:
        for cid, grp in df.groupby(cust_col):
            total  = float(grp[sales_col].sum())
            orders = len(grp)
            region = str(grp[cols['region']].iloc[0]) if cols['region'] else 'Unknown'
            ctype  = str(grp[cols['ctype']].iloc[0]) if cols['ctype'] else ('Returning' if orders>1 else 'New')
            name   = str(grp[cols['name']].iloc[0]) if cols['name'] and cols['name']!=cust_col else str(cid)
            customers[str(cid)] = {
                'id':str(cid),'name':name,'total':total,'orders':orders,
                'avg_order':round(total/orders,2) if orders else 0,
                'region':region,'type':ctype
            }

    cust_list = sorted(customers.values(), key=lambda x: x['total'], reverse=True)
    n = len(cust_list)
    top20 = max(1,int(n*0.2)); mid60 = max(1,int(n*0.6))
    for i,c in enumerate(cust_list):
        c['segment'] = 'High Value' if i<top20 else ('Medium Value' if i<top20+mid60 else 'Low Value')

    products = []
    if cols['product'] and sales_col:
        for pname, grp in df.groupby(cols['product']):
            cat   = str(grp[cols['category']].iloc[0]) if cols['category'] else 'Uncategorized'
            units = int(grp[qty_col].sum()) if qty_col else len(grp)
            rev   = float(grp[sales_col].sum())
            products.append({'name':str(pname),'category':cat,'units':units,'revenue':rev})
        products.sort(key=lambda x: x['revenue'], reverse=True)

    regions = {}
    if cols['region'] and sales_col:
        for r,v in df.groupby(cols['region'])[sales_col].sum().items():
            regions[str(r)] = float(v)

    monthly = {}
    if cols['date'] and sales_col:
        df['_date']  = pd.to_datetime(df[cols['date']], errors='coerce')
        df['_month'] = df['_date'].dt.to_period('M').astype(str)
        for m,v in df.groupby('_month')[sales_col].sum().items():
            if m!='NaT': monthly[str(m)] = float(v)

    categories = {}
    if cols['category'] and sales_col:
        for cat,v in df.groupby(cols['category'])[sales_col].sum().items():
            categories[str(cat)] = float(v)

    total_sales  = float(df[sales_col].sum()) if sales_col else 0
    total_orders = len(df)
    total_cust   = len(cust_list)
    new_cust     = sum(1 for c in cust_list if str(c['type']).strip().lower()=='new')
    ret_cust     = total_cust - new_cust
    high_val     = [c for c in cust_list if c['segment']=='High Value']
    med_val      = [c for c in cust_list if c['segment']=='Medium Value']
    low_val      = [c for c in cust_list if c['segment']=='Low Value']
    avg_order    = total_sales/total_orders if total_orders else 0
    top_region   = max(regions, key=regions.get) if regions else '-'

    kpis = {
        'total_sales':total_sales,'total_sales_fmt':fmtm(total_sales),
        'total_orders':total_orders,'total_customers':total_cust,
        'new_customers':new_cust,'returning_customers':ret_cust,
        'avg_order':avg_order,'avg_order_fmt':fmtm(avg_order),
        'high_value_count':len(high_val),'med_value_count':len(med_val),'low_value_count':len(low_val),
        'top_customer':cust_list[0] if cust_list else None,
        'top_product':products[0] if products else None,
        'top_region':top_region,'total_products':len(products),
    }
    return kpis, cust_list, products, regions, dict(sorted(monthly.items())), categories

# ──────────────────────────────────────────────
#  SUGGESTIONS ENGINE
# ──────────────────────────────────────────────
def get_suggestions(kpis, cust_list, products, regions):
    sug = []
    tc = max(kpis['total_customers'],1)
    ret = kpis['returning_customers']/tc*100
    if ret < 40:
        sug.append({'title':'Improve Customer Retention','priority':'High','icon':'🔄',
            'detail':f'Only {ret:.1f}% customers return. Launch a loyalty/points program, personalized email follow-ups, and exclusive repeat-buyer discounts. Target: 60%+ retention rate.'})
    else:
        sug.append({'title':'Deepen Loyalty Program','priority':'Medium','icon':'⭐',
            'detail':f'{ret:.1f}% retention is strong. Introduce a VIP tier for top 20% spenders with early access, free shipping, and personal account managers to increase CLV further.'})

    hv_pct = kpis['high_value_count']/tc*100
    if hv_pct < 15:
        sug.append({'title':'Upsell to Medium Value Customers','priority':'High','icon':'📈',
            'detail':f'High-value customers are only {hv_pct:.1f}% of base. Run targeted bundle offers, premium upgrade campaigns, and "spend X get Y" promotions to convert medium-value customers.'})

    lv_pct = kpis['low_value_count']/tc*100
    if lv_pct > 40:
        sug.append({'title':'Reactivate Low Value Customers','priority':'Medium','icon':'💌',
            'detail':f'{kpis["low_value_count"]} customers ({lv_pct:.1f}%) are low value. Deploy win-back campaigns with time-limited discounts, product recommendations, and "we miss you" emails.'})

    if products and len(products)>=3:
        top3 = sum(p['revenue'] for p in products[:3])
        tr   = sum(p['revenue'] for p in products)
        if tr>0 and top3/tr>0.6:
            sug.append({'title':'Reduce Product Concentration Risk','priority':'High','icon':'📦',
                'detail':f'Top 3 products generate {top3/tr*100:.0f}% of revenue. Actively market mid-tier products via bundles, cross-sell suggestions, and category landing pages to diversify revenue.'})

    if regions:
        rv = sorted(regions.values(), reverse=True)
        top_r = max(regions, key=regions.get)
        if rv[0] > sum(rv)*0.5:
            sug.append({'title':f'Expand Beyond {top_r}','priority':'Medium','icon':'🗺️',
                'detail':f'{top_r} drives 50%+ of regional revenue. Invest in underperforming regions with targeted digital ads, local partnerships, and region-specific promotions.'})

    aov = kpis['avg_order']
    sug.append({'title':'Increase Average Order Value','priority':'High','icon':'💰',
        'detail':f'Current AOV is {fmtm(aov)}. Add cross-sell recommendations at checkout, free-shipping thresholds, and product bundles to push AOV up by 15–25%.'})

    new_pct = kpis['new_customers']/tc*100
    if new_pct < 25:
        sug.append({'title':'Accelerate New Customer Acquisition','priority':'Medium','icon':'🚀',
            'detail':f'New customers are {new_pct:.1f}% of base. Launch referral programs, paid social campaigns targeting lookalike audiences, and first-time-buyer discount codes.'})

    sug.append({'title':'Run Seasonal & Flash Sales','priority':'Low','icon':'📅',
        'detail':'Use monthly trend data to identify revenue dips. Schedule targeted flash sales, seasonal bundles, and holiday campaigns to maintain consistent monthly revenue through slow periods.'})

    sug.append({'title':'Invest in Customer Data & Analytics','priority':'Low','icon':'📊',
        'detail':'Set up monthly reporting dashboards to track KPIs, segment drift, and product performance in real time. Use data to make faster, evidence-based business decisions.'})

    return sug

# ──────────────────────────────────────────────
#  CHART HELPERS (for PDF)
# ──────────────────────────────────────────────
DARK   = '#0a0d14'
SURF   = '#111520'
SURF2  = '#161c2d'
BLUE   = '#4f8ef7'
PURPLE = '#7c3aed'
GREEN  = '#10d9a0'
AMBER  = '#f59e0b'
RED    = '#ef4444'
CYAN   = '#06b6d4'
TEXT   = '#e8edf8'
MUTED  = '#6b7a9e'
PALETTE = [BLUE,PURPLE,GREEN,AMBER,RED,CYAN,'#84cc16','#ec4899','#a855f7','#14b8a6']

def set_dark_axes(ax, fig):
    fig.patch.set_facecolor(SURF)
    ax.set_facecolor(SURF2)
    ax.tick_params(colors=MUTED, labelsize=8)
    ax.xaxis.label.set_color(MUTED)
    ax.yaxis.label.set_color(MUTED)
    for spine in ax.spines.values():
        spine.set_edgecolor(SURF2)
    ax.grid(axis='y', color=SURF, linewidth=0.8, alpha=0.6)

def chart_to_rl(fig, w_cm=16, h_cm=7):
    buf = BytesIO()
    fig.savefig(buf, format='png', dpi=130, bbox_inches='tight', facecolor=fig.get_facecolor())
    buf.seek(0)
    plt.close(fig)
    return RLImage(buf, width=w_cm*cm, height=h_cm*cm)

def make_sales_trend(monthly):
    if not monthly: return None
    months = list(monthly.keys())
    vals   = list(monthly.values())
    fig, ax = plt.subplots(figsize=(10,4))
    set_dark_axes(ax, fig)
    ax.plot(months, vals, color=BLUE, linewidth=2.5, marker='o', markersize=5, markerfacecolor=BLUE)
    ax.fill_between(months, vals, alpha=0.15, color=BLUE)
    ax.set_title('Monthly Sales Trend', color=TEXT, fontsize=11, pad=10)
    ax.set_xticklabels(months, rotation=45, ha='right', fontsize=7)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x,_: f'${x/1000:.0f}K' if x>=1000 else f'${x:.0f}'))
    return chart_to_rl(fig, 16, 6)

def make_region_bar(regions):
    if not regions: return None
    sreg = dict(sorted(regions.items(), key=lambda x: x[1], reverse=True))
    fig, ax = plt.subplots(figsize=(10,4))
    set_dark_axes(ax, fig)
    bars = ax.bar(list(sreg.keys()), list(sreg.values()),
                  color=PALETTE[:len(sreg)], width=0.55, edgecolor='none', linewidth=0)
    for bar in bars:
        h = bar.get_height()
        ax.text(bar.get_x()+bar.get_width()/2, h+max(sreg.values())*0.01,
                fmtm(h), ha='center', va='bottom', color=TEXT, fontsize=7)
    ax.set_title('Region-wise Revenue', color=TEXT, fontsize=11, pad=10)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x,_: f'${x/1000:.0f}K' if x>=1000 else f'${x:.0f}'))
    return chart_to_rl(fig, 16, 6)

def make_seg_pie(kpis):
    hv = kpis.get('high_value_count',0)
    mv = kpis.get('med_value_count',0)
    lv = kpis.get('low_value_count',0)
    if hv+mv+lv == 0: return None
    fig, ax = plt.subplots(figsize=(6,5))
    fig.patch.set_facecolor(SURF); ax.set_facecolor(SURF)
    wedges, texts, autotexts = ax.pie(
        [hv,mv,lv], labels=['High Value','Medium Value','Low Value'],
        colors=[GREEN,AMBER,RED], autopct='%1.1f%%',
        startangle=90, wedgeprops={'edgecolor':SURF,'linewidth':2})
    for t in texts: t.set_color(MUTED); t.set_fontsize(8)
    for a in autotexts: a.set_color(DARK); a.set_fontsize(8); a.set_fontweight('bold')
    ax.set_title('Customer Segmentation', color=TEXT, fontsize=11, pad=10)
    return chart_to_rl(fig, 8, 7)

def make_new_ret_pie(kpis):
    nc = kpis.get('new_customers',0); rc = kpis.get('returning_customers',0)
    if nc+rc == 0: return None
    fig, ax = plt.subplots(figsize=(6,5))
    fig.patch.set_facecolor(SURF); ax.set_facecolor(SURF)
    wedges, texts, autotexts = ax.pie(
        [nc,rc], labels=['New','Returning'],
        colors=[BLUE,PURPLE], autopct='%1.1f%%',
        startangle=90, wedgeprops={'edgecolor':SURF,'linewidth':2})
    for t in texts: t.set_color(MUTED); t.set_fontsize(8)
    for a in autotexts: a.set_color(TEXT); a.set_fontsize(8); a.set_fontweight('bold')
    ax.set_title('New vs Returning Customers', color=TEXT, fontsize=11, pad=10)
    return chart_to_rl(fig, 8, 7)

def make_top_products(products):
    if not products: return None
    top = products[:10]
    fig, ax = plt.subplots(figsize=(10,5))
    set_dark_axes(ax, fig)
    names = [p['name'][:22] for p in top]
    revs  = [p['revenue'] for p in top]
    bars  = ax.barh(names[::-1], revs[::-1], color=PALETTE[:len(top)], height=0.65, edgecolor='none')
    for bar in bars:
        w = bar.get_width()
        ax.text(w + max(revs)*0.01, bar.get_y()+bar.get_height()/2,
                fmtm(w), va='center', color=TEXT, fontsize=7)
    ax.set_title('Top 10 Products by Revenue', color=TEXT, fontsize=11, pad=10)
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x,_: f'${x/1000:.0f}K' if x>=1000 else f'${x:.0f}'))
    ax.tick_params(axis='y', labelsize=7)
    return chart_to_rl(fig, 16, 7)

def make_category_bar(categories):
    if not categories: return None
    scat = dict(sorted(categories.items(), key=lambda x: x[1], reverse=True))
    fig, ax = plt.subplots(figsize=(10,4))
    set_dark_axes(ax, fig)
    bars = ax.bar(list(scat.keys()), list(scat.values()),
                  color=PALETTE[:len(scat)], width=0.55, edgecolor='none')
    for bar in bars:
        h = bar.get_height()
        ax.text(bar.get_x()+bar.get_width()/2, h+max(scat.values())*0.01,
                fmtm(h), ha='center', va='bottom', color=TEXT, fontsize=7)
    ax.set_title('Revenue by Category', color=TEXT, fontsize=11, pad=10)
    ax.set_xticklabels(list(scat.keys()), rotation=30, ha='right', fontsize=8)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x,_: f'${x/1000:.0f}K' if x>=1000 else f'${x:.0f}'))
    return chart_to_rl(fig, 16, 6)

# ──────────────────────────────────────────────
#  PDF BUILDER
# ──────────────────────────────────────────────
def build_pdf(kpis, cust_list, products, regions, monthly, categories, suggestions, stats):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                             leftMargin=1.8*cm, rightMargin=1.8*cm,
                             topMargin=2*cm, bottomMargin=2*cm)
    W = A4[0] - 3.6*cm

    # ── Styles ──
    h1  = ParagraphStyle('h1',  fontName='Helvetica-Bold', fontSize=22, textColor=colors.HexColor(TEXT),
                          spaceAfter=4, leading=26, alignment=TA_CENTER)
    h2  = ParagraphStyle('h2',  fontName='Helvetica-Bold', fontSize=13, textColor=colors.HexColor(BLUE),
                          spaceBefore=14, spaceAfter=6, leading=16)
    h3  = ParagraphStyle('h3',  fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor(TEXT),
                          spaceBefore=8, spaceAfter=4)
    body= ParagraphStyle('body', fontName='Helvetica',      fontSize=9,  textColor=colors.HexColor(MUTED),
                          leading=14, spaceAfter=4)
    sub = ParagraphStyle('sub',  fontName='Helvetica',      fontSize=8,  textColor=colors.HexColor(MUTED),
                          alignment=TA_CENTER)
    ftr = ParagraphStyle('ftr',  fontName='Helvetica',      fontSize=7,  textColor=colors.HexColor(MUTED),
                          alignment=TA_CENTER, spaceBefore=6)
    val_style = ParagraphStyle('val', fontName='Helvetica-Bold', fontSize=18, leading=22,
                                textColor=colors.HexColor(BLUE), alignment=TA_CENTER)
    lbl_style = ParagraphStyle('lbl', fontName='Helvetica', fontSize=7,
                                textColor=colors.HexColor(MUTED), alignment=TA_CENTER, spaceBefore=2)

    def hr(): return HRFlowable(width='100%', thickness=1, color=colors.HexColor(SURF2), spaceAfter=8, spaceBefore=8)

    def kpi_table(rows):
        """rows = [(label, value), ...]"""
        cells = []
        for label, value in rows:
            cells.append([
                Paragraph(str(value), val_style),
                Paragraph(label, lbl_style)
            ])
        # Split into groups of 3
        tdata = []
        for i in range(0, len(cells), 3):
            group = cells[i:i+3]
            row1 = [g[0] for g in group]
            row2 = [g[1] for g in group]
            while len(row1) < 3: row1.append(Paragraph('', val_style)); row2.append(Paragraph('', lbl_style))
            tdata.append(row1); tdata.append(row2)
        cw = W/3
        t = Table(tdata, colWidths=[cw,cw,cw], rowHeights=None)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor(SURF)),
            ('BOX',        (0,0), (-1,-1), 1, colors.HexColor(SURF2)),
            ('INNERGRID',  (0,0), (-1,-1), 0.5, colors.HexColor(SURF2)),
            ('TOPPADDING', (0,0), (-1,-1), 10),
            ('BOTTOMPADDING',(0,0),(-1,-1), 8),
            ('ALIGN',      (0,0), (-1,-1), 'CENTER'),
            ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
            ('ROUNDEDCORNERS', [4]),
        ]))
        return t

    def data_table(headers, rows, col_widths=None):
        if col_widths is None:
            cw = W / len(headers)
            col_widths = [cw]*len(headers)
        hrow = [Paragraph(h, ParagraphStyle('th', fontName='Helvetica-Bold', fontSize=8,
                           textColor=colors.HexColor(BLUE), alignment=TA_CENTER)) for h in headers]
        tdata = [hrow]
        for i, row in enumerate(rows):
            tdata.append([Paragraph(str(cell), ParagraphStyle('td', fontName='Helvetica', fontSize=8,
                           textColor=colors.HexColor(TEXT if i%2==0 else MUTED), alignment=TA_CENTER))
                          for cell in row])
        t = Table(tdata, colWidths=col_widths, repeatRows=1)
        ts = TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor(SURF2)),
            ('BACKGROUND', (0,1), (-1,-1), colors.HexColor(SURF)),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.HexColor(SURF), colors.HexColor(SURF2)]),
            ('BOX',        (0,0), (-1,-1), 0.5, colors.HexColor(SURF2)),
            ('INNERGRID',  (0,0), (-1,-1), 0.3, colors.HexColor(SURF2)),
            ('TOPPADDING', (0,0), (-1,-1), 6),
            ('BOTTOMPADDING',(0,0),(-1,-1), 6),
            ('ALIGN',      (0,0), (-1,-1), 'CENTER'),
            ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
        ])
        t.setStyle(ts)
        return t

    def seg_badge(seg):
        color_map = {'High Value': GREEN, 'Medium Value': AMBER, 'Low Value': RED}
        c = color_map.get(seg, BLUE)
        return Paragraph(f'<font color="{c}"><b>{seg}</b></font>',
                         ParagraphStyle('badge', fontName='Helvetica-Bold', fontSize=8, alignment=TA_CENTER))

    story = []

    # ── PAGE BACKGROUND ──
    def on_page(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(colors.HexColor(DARK))
        canvas.rect(0, 0, A4[0], A4[1], fill=1, stroke=0)
        canvas.restoreState()

    # ══════════════════════════════════════════
    #  COVER PAGE
    # ══════════════════════════════════════════
    story.append(Spacer(1, 2.5*cm))
    story.append(Paragraph("InsightIQ", ParagraphStyle('brand', fontName='Helvetica-Bold', fontSize=36, leading=42,
                  textColor=colors.HexColor(BLUE), alignment=TA_CENTER, spaceAfter=4)))
    story.append(Paragraph("Business Intelligence Report", ParagraphStyle('sub2', fontName='Helvetica', fontSize=14, leading=18,
                  textColor=colors.HexColor(MUTED), alignment=TA_CENTER, spaceAfter=4)))
    story.append(Spacer(1, 0.4*cm))
    story.append(HRFlowable(width='60%', thickness=2, color=colors.HexColor(BLUE),
                              hAlign='CENTER', spaceAfter=16))
    story.append(Spacer(1, 0.8*cm))

    # Cover KPI highlight
    cover_kpis = [
        ('Total Revenue',     kpis['total_sales_fmt']),
        ('Total Customers',   str(kpis['total_customers'])),
        ('Total Orders',      str(kpis['total_orders'])),
        ('Avg Order Value',   kpis['avg_order_fmt']),
        ('High Value Customers', str(kpis['high_value_count'])),
        ('Total Products',    str(kpis['total_products'])),
    ]
    story.append(kpi_table(cover_kpis))
    story.append(Spacer(1, 1.2*cm))
    story.append(Paragraph(f"Dataset: {stats['rows']:,} rows · {stats['columns']} columns", sub))
    story.append(Spacer(1, 0.4*cm))
    story.append(Paragraph("Designed & Developed by <b>Aditya Jassal</b>", sub))
    story.append(PageBreak())

    # ══════════════════════════════════════════
    #  1. EXECUTIVE SUMMARY
    # ══════════════════════════════════════════
    story.append(Paragraph("1. Executive Summary", h2))
    story.append(hr())
    tc = max(kpis['total_customers'],1)
    ret_rate = kpis['returning_customers']/tc*100
    hv_pct   = kpis['high_value_count']/tc*100
    top_p    = kpis['top_product']['name'] if kpis.get('top_product') else '-'
    top_r    = kpis['top_region']

    exec_text = (
        f"This report presents a comprehensive analysis of the uploaded sales and customer dataset comprising "
        f"<b>{stats['rows']:,} records</b> across {stats['columns']} attributes. "
        f"Total revenue stands at <b>{kpis['total_sales_fmt']}</b> generated from "
        f"<b>{kpis['total_orders']:,} orders</b> by <b>{kpis['total_customers']:,} unique customers</b>. "
        f"The average order value is <b>{kpis['avg_order_fmt']}</b>. "
        f"Customer retention rate is <b>{ret_rate:.1f}%</b> with "
        f"<b>{kpis['returning_customers']}</b> returning and <b>{kpis['new_customers']}</b> new customers. "
        f"<b>{hv_pct:.1f}%</b> of customers fall in the High Value segment. "
        f"The best-performing product is <b>{top_p}</b> and top region is <b>{top_r}</b>."
    )
    story.append(Paragraph(exec_text, body))
    story.append(Spacer(1, 0.5*cm))

    # ══════════════════════════════════════════
    #  2. SALES PERFORMANCE
    # ══════════════════════════════════════════
    story.append(Paragraph("2. Sales Performance", h2))
    story.append(hr())

    sales_chart = make_sales_trend(monthly)
    if sales_chart: story.append(sales_chart)
    story.append(Spacer(1, 0.3*cm))

    if monthly:
        vals = list(monthly.values())
        best_m = max(monthly, key=monthly.get)
        worst_m = min(monthly, key=monthly.get)
        story.append(Paragraph(
            f"Best month: <b>{best_m}</b> ({fmtm(monthly[best_m])}) · "
            f"Weakest month: <b>{worst_m}</b> ({fmtm(monthly[worst_m])}) · "
            f"Monthly avg: <b>{fmtm(sum(vals)/len(vals))}</b>", body))

    story.append(Spacer(1, 0.5*cm))
    region_chart = make_region_bar(regions)
    if region_chart: story.append(region_chart)

    if regions:
        story.append(Spacer(1, 0.3*cm))
        reg_rows = [[r, fmtm(v), f"{v/max(sum(regions.values()),1)*100:.1f}%"]
                    for r,v in sorted(regions.items(), key=lambda x: x[1], reverse=True)]
        story.append(data_table(['Region','Revenue','Share'], reg_rows, [6*cm, 5*cm, 4.2*cm]))

    story.append(PageBreak())

    # ══════════════════════════════════════════
    #  3. CUSTOMER ANALYSIS
    # ══════════════════════════════════════════
    story.append(Paragraph("3. Customer Analysis", h2))
    story.append(hr())

    seg_img = make_seg_pie(kpis)
    ret_img = make_new_ret_pie(kpis)
    if seg_img and ret_img:
        row = [[seg_img, ret_img]]
        t = Table(row, colWidths=[W/2, W/2])
        t.setStyle(TableStyle([('ALIGN',(0,0),(-1,-1),'CENTER'),('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
        story.append(t)
    elif seg_img: story.append(seg_img)

    story.append(Spacer(1, 0.4*cm))
    story.append(Paragraph("3.1 Customer Segments", h3))
    hv = [c for c in cust_list if c['segment']=='High Value']
    mv = [c for c in cust_list if c['segment']=='Medium Value']
    lv = [c for c in cust_list if c['segment']=='Low Value']

    def seg_revenue(seg_list):
        return sum(c['total'] for c in seg_list)

    seg_rows = [
        ['High Value',   len(hv), fmtm(seg_revenue(hv)), fmtm(seg_revenue(hv)/max(len(hv),1)), f"{len(hv)/tc*100:.1f}%"],
        ['Medium Value', len(mv), fmtm(seg_revenue(mv)), fmtm(seg_revenue(mv)/max(len(mv),1)), f"{len(mv)/tc*100:.1f}%"],
        ['Low Value',    len(lv), fmtm(seg_revenue(lv)), fmtm(seg_revenue(lv)/max(len(lv),1)), f"{len(lv)/tc*100:.1f}%"],
    ]
    story.append(data_table(['Segment','Customers','Total Revenue','Avg Spend','% of Base'],
                              seg_rows, [4.5*cm,3.5*cm,4*cm,4*cm,3.2*cm]))

    story.append(Spacer(1, 0.4*cm))
    story.append(Paragraph("3.2 Top 15 Customers", h3))
    top15 = cust_list[:15]
    cust_rows = [[c['name'][:24], fmtm(c['total']), str(c['orders']),
                  fmtm(c['avg_order']), c['region'], c['type'], c['segment']]
                 for c in top15]
    story.append(data_table(['Customer','Total Spent','Orders','Avg Order','Region','Type','Segment'],
                              cust_rows, [4*cm,2.8*cm,2*cm,2.8*cm,2.5*cm,2*cm,3.1*cm]))

    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph("3.3 Bottom 5 Customers (Low Spend)", h3))
    bot5 = cust_list[-5:][::-1]
    bot_rows = [[c['name'][:24], fmtm(c['total']), str(c['orders']), c['region'], c['segment']] for c in bot5]
    story.append(data_table(['Customer','Total Spent','Orders','Region','Segment'],
                              bot_rows, [5.5*cm,3.5*cm,3*cm,3.5*cm,3.7*cm]))

    story.append(PageBreak())

    # ══════════════════════════════════════════
    #  4. PRODUCT PERFORMANCE
    # ══════════════════════════════════════════
    story.append(Paragraph("4. Product Performance", h2))
    story.append(hr())

    prod_chart = make_top_products(products)
    if prod_chart: story.append(prod_chart)
    story.append(Spacer(1, 0.3*cm))

    cat_chart = make_category_bar(categories)
    if cat_chart: story.append(cat_chart)
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("4.1 Top 10 Products", h3))
    top10p = products[:10]
    prod_rows = [[str(i+1), p['name'][:28], p['category'], str(p['units']),
                  fmtm(p['revenue']), fmtm(p['revenue']/max(p['units'],1))]
                 for i,p in enumerate(top10p)]
    story.append(data_table(['#','Product','Category','Units','Revenue','Avg Price'],
                              prod_rows, [1.2*cm,5*cm,3.5*cm,2.5*cm,3*cm,3*cm]))

    if len(products) > 5:
        story.append(Spacer(1, 0.3*cm))
        story.append(Paragraph("4.2 Least Performing Products", h3))
        bot_prod = products[-5:][::-1]
        bp_rows = [[p['name'][:28], p['category'], str(p['units']), fmtm(p['revenue'])] for p in bot_prod]
        story.append(data_table(['Product','Category','Units','Revenue'],
                                  bp_rows, [6*cm,4.5*cm,3.5*cm,5.2*cm]))

    story.append(PageBreak())

    # ══════════════════════════════════════════
    #  5. KEY INSIGHTS
    # ══════════════════════════════════════════
    story.append(Paragraph("5. Key Business Insights", h2))
    story.append(hr())

    insights = []
    tr = kpis['total_sales']
    if monthly:
        vals = list(monthly.values())
        avg_m = sum(vals)/len(vals)
        growth = (vals[-1]-vals[0])/max(vals[0],1)*100 if len(vals)>1 else 0
        insights.append(f"Revenue growth from first to last month: <b>{growth:+.1f}%</b>. Monthly average: <b>{fmtm(avg_m)}</b>.")
    if cust_list:
        top3_rev = sum(c['total'] for c in cust_list[:3])
        insights.append(f"Top 3 customers contribute <b>{top3_rev/max(tr,1)*100:.1f}%</b> of total revenue — concentration risk should be monitored.")
    if products:
        top3p = sum(p['revenue'] for p in products[:3])
        insights.append(f"Top 3 products drive <b>{top3p/max(tr,1)*100:.1f}%</b> of product revenue. Diversify by promoting mid-tier SKUs.")
    if regions:
        top_reg_rev = max(regions.values())
        insights.append(f"Top region contributes <b>{top_reg_rev/max(tr,1)*100:.1f}%</b> of revenue. Underperforming regions present expansion opportunities.")
    insights.append(f"Average Order Value is <b>{kpis['avg_order_fmt']}</b>. Cross-sell and bundling can improve this by 15–25%.")
    insights.append(f"<b>{kpis['high_value_count']}</b> customers ({kpis['high_value_count']/tc*100:.1f}%) are High Value, generating disproportionate revenue — protect these relationships aggressively.")
    insights.append(f"Customer retention rate: <b>{ret_rate:.1f}%</b>. {'Requires improvement — target 60%+.' if ret_rate<40 else 'Healthy — focus on converting new customers to loyal ones.'}")

    for i, ins in enumerate(insights, 1):
        story.append(Paragraph(f"{i}. {ins}", body))
    story.append(Spacer(1, 0.5*cm))

    # ══════════════════════════════════════════
    #  6. BUSINESS SUGGESTIONS
    # ══════════════════════════════════════════
    story.append(Paragraph("6. Business Recommendations", h2))
    story.append(hr())

    priority_colors = {'High': RED, 'Medium': AMBER, 'Low': GREEN}
    for i, sug in enumerate(suggestions, 1):
        pc = priority_colors.get(sug.get('priority','Low'), GREEN)
        story.append(KeepTogether([
            Paragraph(f"{i}. {sug.get('icon','')}  {sug['title']}  "
                      f"[<font color='{pc}'><b>{sug.get('priority','')}</b></font>]", h3),
            Paragraph(sug['detail'], body),
            Spacer(1, 0.2*cm)
        ]))

    story.append(PageBreak())

    # ══════════════════════════════════════════
    #  7. DATA QUALITY
    # ══════════════════════════════════════════
    story.append(Paragraph("7. Dataset Overview & Data Quality", h2))
    story.append(hr())

    sum_rows = [
        ['Total Rows',       f"{stats['rows']:,}"],
        ['Total Columns',    str(stats['columns'])],
        ['Sales Column',     stats.get('col_map',{}).get('sales','-') or '-'],
        ['Customer Column',  stats.get('col_map',{}).get('customer','-') or '-'],
        ['Product Column',   stats.get('col_map',{}).get('product','-') or '-'],
        ['Region Column',    stats.get('col_map',{}).get('region','-') or '-'],
        ['Date Column',      stats.get('col_map',{}).get('date','-') or '-'],
    ]
    story.append(data_table(['Attribute','Value'], sum_rows, [8*cm, 11.2*cm]))
    story.append(Spacer(1, 0.4*cm))
    story.append(Paragraph("Column Details", h3))
    col_rows = [[c, str(stats['dtypes'].get(c,'?')), str(stats['null_counts'].get(c,0))]
                for c in stats['column_names'][:25]]
    story.append(data_table(['Column','Data Type','Null Count'], col_rows, [8*cm,5*cm,6.2*cm]))

    # ── Footer ──
    story.append(Spacer(1, 1*cm))
    story.append(HRFlowable(width='100%', thickness=0.5, color=colors.HexColor(SURF2)))
    story.append(Paragraph(
        "InsightIQ — Customer & Sales Analytics Report &nbsp;·&nbsp; "
        "Designed & Developed by <b>Aditya Jassal</b> &nbsp;·&nbsp; "
        "Powered by Python &amp; ReportLab", ftr))

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    buf.seek(0)
    return buf

# ──────────────────────────────────────────────
#  DATASET CLEANER
# ──────────────────────────────────────────────
def clean_dataframe(df, cols):
    c = df.copy()
    # Strip strings
    for col in c.select_dtypes(include='object').columns:
        c[col] = c[col].astype(str).str.strip()
        c[col] = c[col].replace({'nan':None,'None':None,'NaN':None,'':None})
    # Numeric coercion
    if cols['sales']:
        c[cols['sales']] = pd.to_numeric(c[cols['sales']].astype(str).str.replace(r'[^\d.]','',regex=True),errors='coerce')
    if cols['quantity']:
        c[cols['quantity']] = pd.to_numeric(c[cols['quantity']], errors='coerce')
    if cols['profit']:
        c[cols['profit']] = pd.to_numeric(c[cols['profit']].astype(str).str.replace(r'[^\d.\-]','',regex=True),errors='coerce')
    # Date normalize
    if cols['date']:
        c[cols['date']] = pd.to_datetime(c[cols['date']], errors='coerce')
        c['year']       = c[cols['date']].dt.year.astype('Int64')
        c['month']      = c[cols['date']].dt.month.astype('Int64')
        c['month_name'] = c[cols['date']].dt.strftime('%B')
        c[cols['date']] = c[cols['date']].dt.strftime('%Y-%m-%d')
    # Fill nulls
    for col in c.select_dtypes(include='number').columns:
        c[col] = c[col].fillna(0)
    for col in c.select_dtypes(include='object').columns:
        c[col] = c[col].fillna('Unknown')
    # Remove duplicates
    c = c.drop_duplicates().reset_index(drop=True)
    # Add segment column
    cust_col = cols['customer'] or cols['name']
    sales_col = cols['sales']
    if cust_col and sales_col:
        spend = c.groupby(cust_col)[sales_col].sum().reset_index()
        spend.columns = [cust_col,'_ts']
        spend = spend.sort_values('_ts',ascending=False).reset_index(drop=True)
        n = len(spend); top20 = max(1,int(n*0.2)); mid60 = max(1,int(n*0.6))
        def seg(i): return 'High Value' if i<top20 else ('Medium Value' if i<top20+mid60 else 'Low Value')
        spend['customer_segment'] = [seg(i) for i in range(n)]
        c = c.merge(spend[[cust_col,'customer_segment']], on=cust_col, how='left')
    return c

# ──────────────────────────────────────────────
#  ROUTES
# ──────────────────────────────────────────────
@app.route('/')
def index():
    frontend_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'frontend'))
    return send_file(os.path.join(frontend_dir, 'index.html'))

@app.route('/health')
def health():
    return jsonify({'status':'running'})

@app.route('/analyze', methods=['POST'])
def analyze():
    if 'file' not in request.files:
        return jsonify({'error':'No file uploaded'}), 400
    try:
        file = request.files['file']
        df   = load_file(file)
        cols = detect_cols(df)
        kpis, cust_list, products, regions, monthly, categories = process_dataframe(df.copy(), cols)
        suggestions = get_suggestions(kpis, cust_list, products, regions)
        stats = {
            'rows':len(df),'columns':len(df.columns),
            'column_names':df.columns.tolist(),
            'null_counts':{c:int(v) for c,v in df.isnull().sum().items()},
            'dtypes':{c:str(t) for c,t in df.dtypes.items()},
            'col_map':{k:v for k,v in cols.items() if v}
        }
        preview = df.head(10).fillna('').astype(str).to_dict(orient='records')
        return jsonify({
            'success':True, 'stats':stats, 'kpis':kpis,
            'customers':cust_list[:100], 'products':products[:50],
            'regions':regions, 'monthly':monthly,
            'categories':categories, 'suggestions':suggestions, 'preview':preview,
        })
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error':str(e)}), 500

@app.route('/report', methods=['POST'])
def report():
    if 'file' not in request.files:
        return jsonify({'error':'No file'}), 400
    try:
        file = request.files['file']
        df   = load_file(file)
        cols = detect_cols(df)
        kpis, cust_list, products, regions, monthly, categories = process_dataframe(df.copy(), cols)
        suggestions = get_suggestions(kpis, cust_list, products, regions)
        stats = {
            'rows':len(df),'columns':len(df.columns),
            'column_names':df.columns.tolist(),
            'null_counts':{c:int(v) for c,v in df.isnull().sum().items()},
            'dtypes':{c:str(t) for c,t in df.dtypes.items()},
            'col_map':{k:v for k,v in cols.items() if v}
        }
        pdf_buf = build_pdf(kpis, cust_list, products, regions, monthly, categories, suggestions, stats)
        return send_file(pdf_buf, mimetype='application/pdf',
                         as_attachment=True, download_name='InsightIQ_Report.pdf')
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error':str(e)}), 500

@app.route('/clean', methods=['POST'])
def clean():
    if 'file' not in request.files:
        return jsonify({'error':'No file'}), 400
    try:
        file    = request.files['file']
        fmt_out = request.form.get('format','csv')
        df      = load_file(file)
        cols    = detect_cols(df)
        cleaned = clean_dataframe(df, cols)
        buf = io.BytesIO()
        if fmt_out == 'xlsx':
            cleaned.to_excel(buf, index=False, engine='openpyxl')
            mime  = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            fname = 'InsightIQ_Cleaned.xlsx'
        else:
            cleaned.to_csv(buf, index=False)
            mime  = 'text/csv'
            fname = 'InsightIQ_Cleaned.csv'
        buf.seek(0)
        return send_file(buf, mimetype=mime, as_attachment=True, download_name=fname)
    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error':str(e)}), 500

if __name__ == '__main__':
    print("InsightIQ Backend running on http://localhost:5050")
    app.run(debug=False, port=5050, host='0.0.0.0')
