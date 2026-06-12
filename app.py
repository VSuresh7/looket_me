import os
import sys
import json
import uuid
import random
import urllib.parse
from datetime import datetime

import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, session
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from twilio.rest import Client

# Helper function for read-only UI assets (HTML/CSS)
def get_resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# Initialize Flask with bundled asset paths
app = Flask(__name__, 
            template_folder=get_resource_path("templates"),
            static_folder=get_resource_path("static"))

# Setup a writable user path in the Windows Documents folder for your Excel database
DOCUMENTS_DIR = os.path.join(os.path.expanduser("~"), "Documents", "LooketMe_Data")
if not os.path.exists(DOCUMENTS_DIR):
    os.makedirs(DOCUMENTS_DIR)

USERS_FILE = os.path.join(DOCUMENTS_DIR, "users.xlsx")

# Automatically generate the empty Excel sheet structure if running for the first time
if not os.path.exists(USERS_FILE):
    df_empty = pd.DataFrame(columns=["id", "username", "email", "password"])
    df_empty.to_excel(USERS_FILE, index=False, engine="openpyxl")
# ====================================================================
# PYINSTALLER PATH RESOLUTION HELPER
# ====================================================================
def get_resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# Resolve the UI templates and static folders dynamically
template_dir = get_resource_path('templates')
static_dir = get_resource_path('static')

# Initialize the Flask application ONCE with proper asset paths
app = Flask(__name__, template_folder=template_dir, static_folder=static_dir)
app.secret_key = "LOOKET_SUPER_SECRET_KEY_BRO"

# ====================================================================
# PERMANENT DATA STORAGE PATHS (Prevents login/signup data loss)
# ====================================================================
def get_permanent_data_path(filename):
    """ Ensures database files live permanently inside the user's Documents folder """
    data_dir = os.path.join(os.path.expanduser("~"), "Documents", "LooketMe_Data")
    if not os.path.exists(data_dir):
        os.makedirs(data_dir, exist_ok=True)
    return os.path.join(data_dir, filename)

PRODUCTS_FILE = get_permanent_data_path("products.xlsx")
SALES_FILE = get_permanent_data_path("sales.xlsx")
CUSTOMERS_FILE = get_permanent_data_path("customers.xlsx")
USERS_FILE = get_permanent_data_path("users.xlsx")

# Seed default admin account if users spreadsheet doesn't exist
if not os.path.exists(USERS_FILE):
    pd.DataFrame([{
        "id": "U-1", 
        "username": "admin", 
        "password": generate_password_hash("admin123")
    }]).to_excel(USERS_FILE, index=False, engine="openpyxl")

# ====================================================================
# TWILIO OTP GATEWAY INTEGRATION CONFIGURATION
# ====================================================================
TWILIO_ACCOUNT_SID = 'your_account_sid_here'
TWILIO_AUTH_TOKEN = 'your_auth_token_here'
TWILIO_PHONE_NUMBER = '+917997941247' 
SANDBOX_MODE = True 

def send_otp_sms(target_phone, otp_code):
    if SANDBOX_MODE:
        print("\n" + "="*60)
        print(f"🔥 LOOKET ME DEV MODE OTP: {otp_code} (Target: {target_phone})")
        print("="*60 + "\n")
        return True
    try:
        twilio_client = Client(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN)
        twilio_client.messages.create(
            body=f"LOOKET ME: Your verification OTP code is {otp_code}. Valid for 5 minutes.",
            from_=TWILIO_PHONE_NUMBER,
            to=target_phone
        )
        return True
    except Exception as e:
        print(f"SMS Gateway dispatching failure: {e}")
        return False

# ====================================================================
# SECURITY APP SETUP
# ====================================================================
ENABLE_AUTH = True  
STORE_CONFIG = {"owner": "Looket Executive", "upi": "looketstore@upi"}

login_manager = LoginManager(app)
login_manager.login_view = "login"
login_manager.login_message = "Please sign in to access the system."

class User(UserMixin):
    def __init__(self, id, username):
        self.id = id
        self.username = username

@login_manager.user_loader
def load_user(uid):
    if not os.path.exists(USERS_FILE):
        return None
    df = pd.read_excel(USERS_FILE, engine="openpyxl")
    row = df[df["id"].astype(str) == str(uid)]
    return User(str(row["id"].values[0]), str(row["username"].values[0])) if not row.empty else None

def cond_login_required(f):
    return login_required(f) if ENABLE_AUTH else f

def normalize_barcode(val):
    if pd.isna(val): 
        return ""
    if isinstance(val, (int, float)):
        if val == int(val):
            return str(int(val))
        return str(val)
    val_str = str(val).strip()
    if val_str.endswith('.0'):
        return val_str[:-2]
    if '.' in val_str:
        parts = val_str.split('.')
        if parts[1] == '0' or parts[1] == '': 
            return parts[0]
    return val_str

@app.context_processor
def inject_global_warnings():
    alert_banner = None
    try:
        if os.path.exists(PRODUCTS_FILE):
            df = pd.read_excel(PRODUCTS_FILE, engine="openpyxl")
            if not df.empty:
                df.columns = [str(c).strip().lower() for c in df.columns]
                size_cols = ["xs", "s", "m", "l", "xl", "xxl"]
                for c in size_cols:
                    if c in df.columns:
                        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0).astype(int)
                total_units = df[[c for c in size_cols if c in df.columns]].sum().sum()
                if total_units < 50:
                    alert_banner = f"🚨 CRITICAL ALERT: Store inventory is running low ({total_units} pcs remaining). Import stock immediately!"
    except Exception:
        pass
    return dict(global_stock_alert=alert_banner, shop_owner=STORE_CONFIG["owner"], upi_id=STORE_CONFIG["upi"])

# ====================================================================
# ROUTES & API ENDPOINTS
# ====================================================================
@app.route("/")
@cond_login_required
def dashboard():
    counts = {
        "Shirt": 0, "T-Shirt": 0, "Jeans": 0, "Trouser": 0, "Kurta": 0, "Track Pant": 0,
        "Kid's Shirt": 0, "Kid's Jeans": 0, "Kid's Trouser": 0, "Kid's Kurta": 0, 
        "Jacket": 0, "Kid's Jacket": 0
    }
    t_purch, t_curr, t_inv, t_rev, low_s = 0, 0, 0.0, 0.0, 0
    matrix = {}
    
    if os.path.exists(PRODUCTS_FILE):
        df = pd.read_excel(PRODUCTS_FILE, engine="openpyxl")
        if not df.empty:
            df.columns = [str(c).strip().lower() for c in df.columns]
            cols = ["xs", "s", "m", "l", "xl", "xxl"]
            df["barcode"] = df["barcode"].apply(normalize_barcode).astype(str)
            
            for c in cols: 
                if c in df.columns:
                    df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0).astype(int)
                    
            df["total_quantity"] = df[[c for c in cols if c in df.columns]].sum(axis=1)
            t_curr = int(df["total_quantity"].sum())
            low_s = len(df[df["total_quantity"] <= 10])
            
            for _, r in df.iterrows():
                b, cat, qty, pr = str(r["barcode"]), str(r["category"]).lower(), int(r["total_quantity"]), float(r["purchase_price"])
                avail = [c.upper() for c in cols if c in df.columns and int(r[c]) > 0]
                matrix[f"{r['product_name']} [{b}]"] = ", ".join([f"{s}({r[s.lower()]}pcs)" for s in avail]) if avail else "Out of Stock"
                
                if "tshirt" in cat or "t-shirt" in cat: 
                    counts["T-Shirt"] += qty
                elif "kid's shirt" in cat or "kids shirt" in cat: 
                    counts["Kid's Shirt"] += qty
                elif "shirt" in cat: 
                    counts["Shirt"] += qty
                elif "kid's jean" in cat or "kids jean" in cat: 
                    counts["Kid's Jeans"] += qty
                elif "jean" in cat: 
                    counts["Jeans"] += qty
                elif "kid's trouser" in cat or "kids trouser" in cat: 
                    counts["Kid's Trouser"] += qty
                elif "trouser" in cat: 
                    counts["Trouser"] += qty
                elif "kid's kurta" in cat or "kids kurta" in cat: 
                    counts["Kid's Kurta"] += qty
                elif "kurta" in cat: 
                    counts["Kurta"] += qty
                elif "kid's jacket" in cat or "kids jacket" in cat: 
                    counts["Kid's Jacket"] += qty
                elif "jacket" in cat: 
                    counts["Jacket"] += qty
                else: 
                    counts["Track Pant"] += qty
                    
                t_inv += (qty * pr)
                
    if os.path.exists(SALES_FILE):
        try:
            sdf = pd.read_excel(SALES_FILE, engine="openpyxl")
            if not sdf.empty:
                t_rev = float(pd.to_numeric(sdf["total_price"], errors='coerce').fillna(0.0).sum())
                t_purch = t_curr + int(pd.to_numeric(sdf["quantity_sold"], errors='coerce').fillna(0).sum())
        except Exception:
            pass
            
    return render_template("dashboard.html", category_counts=counts, total_purchased_stock=t_purch, total_current_stock=t_curr, total_investment_value=round(t_inv,2), total_sales_revenue=round(t_rev,2), low_stock=low_s, size_avail_matrix=matrix)

@app.route("/api/get_product")
@cond_login_required
def api_get_product():
    barcode = str(request.args.get("barcode", "")).strip()
    if not barcode:
        return jsonify({"error": "Empty Barcode Input"})

    search_target = normalize_barcode(barcode)
    if os.path.exists(PRODUCTS_FILE):
        try:
            df = pd.read_excel(PRODUCTS_FILE, engine="openpyxl")
            if not df.empty:
                df.columns = [str(c).strip().lower() for c in df.columns]
                df["barcode_clean"] = df["barcode"].apply(normalize_barcode).astype(str)

                row = df[df["barcode_clean"] == search_target]
                if not row.empty:
                    res = row.iloc[0].to_dict()
                    clean_response = {}
                    for k, v in res.items():
                        if k in ["xs", "s", "m", "l", "xl", "xxl"]:
                            if isinstance(v, (int, float)):
                                clean_response[k.upper()] = int(v) if not pd.isna(v) else 0
                            else:
                                clean_response[k.upper()] = int(pd.to_numeric(v, errors='coerce').fillna(0))
                        elif k.startswith("price_") or k == "purchase_price":
                            if isinstance(v, (int, float)):
                                clean_response[k] = float(v) if not pd.isna(v) else 0.0
                            else:
                                clean_response[k] = float(pd.to_numeric(v, errors='coerce').fillna(0.0))
                        else:
                            clean_response[k] = str(v)
                    
                    clean_response["size_prices"] = {
                        "XS": float(pd.to_numeric(res.get("price_xs", 0.0), errors='coerce').fillna(0.0)) if not isinstance(res.get("price_xs", 0.0), (int, float)) else float(res.get("price_xs", 0.0)),
                        "S": float(pd.to_numeric(res.get("price_s", 0.0), errors='coerce').fillna(0.0)) if not isinstance(res.get("price_s", 0.0), (int, float)) else float(res.get("price_s", 0.0)),
                        "M": float(pd.to_numeric(res.get("price_m", 0.0), errors='coerce').fillna(0.0)) if not isinstance(res.get("price_m", 0.0), (int, float)) else float(res.get("price_m", 0.0)),
                        "L": float(pd.to_numeric(res.get("price_l", 0.0), errors='coerce').fillna(0.0)) if not isinstance(res.get("price_l", 0.0), (int, float)) else float(res.get("price_l", 0.0)),
                        "XL": float(pd.to_numeric(res.get("price_xl", 0.0), errors='coerce').fillna(0.0)) if not isinstance(res.get("price_xl", 0.0), (int, float)) else float(res.get("price_xl", 0.0)),
                        "XXL": float(pd.to_numeric(res.get("price_xxl", 0.0), errors='coerce').fillna(0.0)) if not isinstance(res.get("price_xxl", 0.0), (int, float)) else float(res.get("price_xxl", 0.0))
                    }
                    return jsonify(clean_response)
        except Exception as e:
            print("API Exception:", e)
            
    return jsonify({"error": "Not Found"})

@app.route("/add_product", methods=["POST"])
@cond_login_required
def add_product():
    try:
        df = pd.read_excel(PRODUCTS_FILE, engine="openpyxl")
    except Exception:
        df = pd.DataFrame()

    b = str(request.form["barcode"]).strip()
    s = str(request.form["incoming_size"]).strip().lower()
    q = int(request.form["incoming_qty"])
    sale_p = float(request.form["sale_price"])
    price_col = f"price_{s}"

    if not df.empty:
        df.columns = [str(c).strip().lower() for c in df.columns]
        df["barcode_temp"] = df["barcode"].apply(normalize_barcode).astype(str)
        search_target = str(normalize_barcode(b))
        match_mask = df["barcode_temp"] == search_target
    else:
        match_mask = pd.Series([False])

    if not df.empty and match_mask.any():
        idx = df[match_mask].index[0]
        current_val = df.at[idx, s] if s in df.columns else 0
        current_stock = 0 if pd.isna(current_val) or str(current_val).strip() == "" else int(pd.to_numeric(current_val, errors='coerce') or 0)
            
        df.at[idx, s] = current_stock + q
        df.at[idx, price_col] = sale_p 
        
        size_cols = ["xs","s","m","l","xl","xxl"]
        total_pieces = 0
        for c in size_cols:
            if c in df.columns:
                cell_val = df.at[idx, c]
                total_pieces += 0 if pd.isna(cell_val) or str(cell_val).strip() == "" else int(pd.to_numeric(cell_val, errors='coerce') or 0)
        
        df.at[idx, "total_quantity"] = total_pieces
        if "barcode_temp" in df.columns:
            df = df.drop(columns=["barcode_temp"])
    else:
        nr = {
            "barcode": b, "product_name": request.form["product_name"], "category": request.form["category"], 
            "purchase_price": float(request.form["purchase_price"]), 
            "xs":0,"s":0,"m":0,"l":0,"xl":0,"xxl":0,
            "price_xs":0.0,"price_s":0.0,"price_m":0.0,"price_l":0.0,"price_xl":0.0,"price_xxl":0.0,
            "total_quantity": q
        }
        nr[s] = q
        nr[price_col] = sale_p
        
        if not df.empty and "barcode_temp" in df.columns:
            df = df.drop(columns=["barcode_temp"])
            
        df = pd.concat([df, pd.DataFrame([nr])], ignore_index=True)

    final_headers = []
    for col in df.columns:
        if col in ["barcode","product_name","category","purchase_price","total_quantity"]:
            final_headers.append(col)
        elif col in ["xs","s","m","l","xl","xxl"]:
            final_headers.append(col.upper())
        elif col.startswith("price_"):
            parts = col.split("_")
            final_headers.append(f"price_{parts[1].upper()}")
        else:
            final_headers.append(col)
            
    df.columns = final_headers
    df.to_excel(PRODUCTS_FILE, index=False, engine="openpyxl")
    return redirect(url_for("inventory"))

@app.route("/inventory")
@cond_login_required
def inventory():
    try:
        df = pd.read_excel(PRODUCTS_FILE, engine="openpyxl")
        df["barcode"] = df["barcode"].apply(normalize_barcode).astype(str)
        p = df.to_dict(orient="records")
    except Exception: p = []
    return render_template("inventory.html", products=p)

@app.route("/billing", methods=["GET", "POST"])
@cond_login_required
def billing():
    receipt_data, whatsapp_url, error_msg = None, None, None
    if request.method == "POST":
        try:
            cname = request.form.get("customer_name", "Customer").strip()
            cphone = request.form.get("customer_phone", "").strip()
            raw_cart_data = request.form.get("cart_data", "").strip()
            
            if not raw_cart_data or raw_cart_data == "[]":
                return render_template("billing.html", error_msg="Checkout Error: Your shopping cart is empty.")

            cart_items = json.loads(raw_cart_data)
            pdf = pd.read_excel(PRODUCTS_FILE, engine="openpyxl")
            
            original_headers = list(pdf.columns)
            pdf.columns = [str(c).strip().lower() for c in pdf.columns]
            pdf["barcode_clean"] = pdf["barcode"].apply(normalize_barcode).astype(str)

            processed_items = []
            new_sales_batch = [] 
            gross_total = 0.0
            total_discount = 0.0
            final_payable = 0.0
            
            tx_id = "INV-" + str(uuid.uuid4())[:6].upper()
            current_timestamp = datetime.now().strftime("%d-%m-%Y %I:%M:%S %p")

            for item in cart_items:
                b_target = str(item.get("barcode", "")).strip()
                s_target = str(item.get("size_sold", "")).strip().lower()
                qty_target = int(item.get("quantity", 1))
                p_name_target = str(item.get("product_name", "Apparel Item"))
                
                search_target = str(normalize_barcode(b_target))
                row_mask = pdf["barcode_clean"] == search_target
                
                rate = float(item.get("rate", 0.0))
                line_discount = float(item.get("discount", 0.0))
                line_final = float(item.get("final_price", (rate * qty_target) - line_discount))

                if row_mask.any():
                    idx = pdf[row_mask].index[0]
                    if s_target in pdf.columns:
                        cell_val = pdf.at[idx, s_target]
                        avail_q = 0 if pd.isna(cell_val) else int(cell_val)
                        
                        if avail_q < qty_target:
                            qty_target = max(0, avail_q)
                        
                        if qty_target > 0:
                            pdf.at[idx, s_target] = avail_q - qty_target
                        
                        size_cols = ["xs","s","m","l","xl","xxl"]
                        total_pieces = 0
                        for c in size_cols:
                            if c in pdf.columns:
                                v_cell = pdf.at[idx, c]
                                if not pd.isna(v_cell):
                                    total_pieces += int(v_cell)
                        pdf.at[idx, "total_quantity"] = total_pieces

                gross_total += (rate * qty_target)
                total_discount += line_discount
                final_payable += line_final
                
                processed_items.append({
                    "product_name": p_name_target,
                    "size_sold": s_target.upper(),
                    "quantity": qty_target,
                    "rate": round(rate, 2),
                    "final_price": round(line_final, 2)
                })

                new_sales_batch.append({
                    "transaction_id": tx_id, "date": current_timestamp, "customer_name": cname, "customer_phone": cphone,
                    "barcode": b_target, "product_name": p_name_target, "size_sold": s_target.upper(), "quantity_sold": qty_target,
                    "total_price": round(line_final, 2), "discount_applied": round(line_discount, 2)
                })

            try:
                sdf = pd.read_excel(SALES_FILE, engine="openpyxl") if os.path.exists(SALES_FILE) else pd.DataFrame()
                sdf = pd.concat([sdf, pd.DataFrame(new_sales_batch)], ignore_index=True)
                sdf.to_excel(SALES_FILE, index=False, engine="openpyxl")
            except Exception as sales_err:
                print(f"Sales Database File Write Lock Exception: {sales_err}")

            if not pdf.empty:
                if "barcode_clean" in pdf.columns:
                    pdf = pdf.drop(columns=["barcode_clean"])
                restored_headers = []
                for col in pdf.columns:
                    matched_orig = [orig for orig in original_headers if orig.lower() == col]
                    restored_headers.append(matched_orig[0] if matched_orig else col)
                pdf.columns = restored_headers
                pdf.to_excel(PRODUCTS_FILE, index=False, engine="openpyxl")

            receipt_data = {
                "invoice_id": str(tx_id), "date": str(current_timestamp), "customer_name": str(cname), "customer_phone": str(cphone),
                "items": processed_items, "gross_total": str(round(gross_total, 2)), "total_discount": str(round(total_discount, 2)),
                "final_payable": str(round(final_payable, 2))
            }

            msg = f"   LOOKET MENS WEAR \n---------------------------------\n"
            msg += f"Hello {cname}, thank you for shopping with us! \nHere is your e-receipt:-\n\n"
            msg += f"RETAIL INVOICE !!:\nBill No: {tx_id}\nDate: {current_timestamp}\n---------------------------------\n|Sl. | Item Name | Qty | Rate | Amount|\n"
            
            for idx, item in enumerate(processed_items, 1):
                msg += f"|{idx}. {item['product_name']}-{item['size_sold']}|\n      {item['quantity']}  x  ₹{item['rate']}  =  ₹{item['final_price']}\n"
                
            msg += f"---------------------------------\nTotal Qty:{sum(int(i['quantity']) for i in processed_items)}\n"
            msg += f"Gross Total: ₹{round(gross_total, 2)}\nDiscount: ₹{round(total_discount, 2)}\n---------------------------------\n"
            msg += f"TOTAL PAID: ₹{receipt_data['final_payable']}*\n---------------------------------\n Customer Name: {cname}\n Mobile No: {cphone}\n\n"
            msg += f"Save Paper, Promote e-Bill \n --VISIT AGAIN-- "

            whatsapp_url = f"https://api.whatsapp.com/send?phone={cphone}&text={urllib.parse.quote(msg)}"
        except Exception as e:
            error_msg = f"System Error processing checkout parameters: {e}"
            print("CRITICAL BILLING EXCEPTION LOG:", e)
            
    return render_template("billing.html", receipt_data=receipt_data, whatsapp_url=whatsapp_url, error_msg=error_msg)

# ====================================================================
# SYSTEM AUTHENTICATION MODES (SIGNUP, LOGIN, VERIFY OTP)
# ====================================================================
# 2. Updated and protected Signup/Registration Route
@app.route("/signup", methods=["GET", "POST"]) # Change to "/register" if your HTML form action calls /register
def signup():
    if request.method == "POST":
        # Extract inputs safely from your signup template forms
        email = request.form.get("email", "").strip()
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()
        
        # Fallback if fields are missing in HTML
        if not username:
            username = email.split('@')[0] # Auto-create username from email if blank
            
        if not email or not password:
            flash("Email and Password are required fields.")
            return render_template("signup.html")
            
        try:
            # Read current users
            df = pd.read_excel(USERS_FILE, engine="openpyxl")
            
            # Ensure columns exist dynamically to prevent append errors
            for col in ["id", "username", "email", "password"]:
                if col not in df.columns:
                    df[col] = None
            
            # Check if user already exists
            if email in df["email"].astype(str).values:
                flash("Email already registered, bro.")
                return render_template("signup.html")
                
            # Create a secure hashed password and registration dict
            hashed_p = generate_password_hash(password)
            user_id = str(len(df) + 1)
            
            new_user = {
                "id": user_id,
                "username": username,
                "email": email,
                "password": hashed_p
            }
            
            # Append new user row to the dataframe
            df = pd.concat([df, pd.DataFrame([new_user])], ignore_index=True)
            
            # Save data back safely down to your documents sheet
            df.to_excel(USERS_FILE, index=False, engine="openpyxl")
            
            flash("Account created successfully! Please sign in.")
            return redirect(url_for("login"))
            
        except Exception as e:
            # If it fails, this print tells you exactly why in your terminal!
            print(f"❌ DATABASE REGISTRATION CRASH ERROR: {e}")
            flash("Server error writing database record.")
            return render_template("signup.html")

    return render_template("signup.html")

@app.route("/verify_otp", methods=["GET", "POST"])
def verify_otp():
    if request.method == "POST":
        user_otp_attempt = request.form.get("otp_attempt", "").strip()
        
        if user_otp_attempt == session.get("current_otp"):
            temp_data = session.get("temp_user")
            if not temp_data:
                flash("Session tracking window closed. Restart registration process.")
                return redirect(url_for("signup"))
            
            df = pd.read_excel(USERS_FILE, engine="openpyxl")
            new_user_id = f"U-{len(df) + 1}"
            new_user_row = {
                "id": new_user_id,
                "username": temp_data["username"],
                "password": temp_data["password"]
            }
            
            df = pd.concat([df, pd.DataFrame([new_user_row])], ignore_index=True)
            df.to_excel(USERS_FILE, index=False, engine="openpyxl")
            
            session.pop("current_otp", None)
            session.pop("temp_user", None)
            
            login_user(User(new_user_id, temp_data["username"]))
            return redirect(url_for("dashboard"))
        else:
            flash("Incorrect OTP entered, bro. Try again.")
            
    return render_template("verify_otp.html")
                           
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        # Match 'email' with the name="email" attribute from login.html
        u = request.form.get("email", "").strip()
        p = request.form.get("password", "").strip()
        
        if not os.path.exists(USERS_FILE):
            flash("User database missing.")
            return render_template("login.html")
            
        # Read the Excel user database
        df = pd.read_excel(USERS_FILE, engine="openpyxl")
        
        # Clean string formats to eliminate spaces
        if "username" in df.columns:
            df["username"] = df["username"].astype(str).str.strip()
        if "email" in df.columns:
            df["email"] = df["email"].astype(str).str.strip()
        
        # Check if the input matches either the email column OR the username column
        if "email" in df.columns and "username" in df.columns:
            r = df[(df["email"] == u) | (df["username"] == u)]
        elif "email" in df.columns:
            r = df[df["email"] == u]
        else:
            r = df[df["username"] == u]
        
        # Verify user matches and validate password hash
        if not r.empty:
            hashed_password = str(r.iloc[0]["password"])
            if check_password_hash(hashed_password, p):
                user_id = str(r.iloc[0]["id"])
                # Fallback to email if username column doesn't exist
                user_name = str(r.iloc[0]["username"]) if "username" in df.columns else str(r.iloc[0]["email"])
                
                login_user(User(user_id, user_name))
                return redirect(url_for("dashboard"))
        
        # Fallback error message if authentication fails
        flash("Invalid Credentials, bro.")
        
    return render_template("login.html")

@app.route("/logout")
@cond_login_required
def logout():
    logout_user()
    return redirect(url_for("login"))

# ====================================================================
# REPORTS & CRM
# ====================================================================
@app.route("/reports")
@cond_login_required
def reports():
    t_sl, m_sl, spends, profit = 0.0, 0.0, 0.0, 0.0
    cur_d, cur_m = datetime.now().strftime("%Y-%m-%d"), datetime.now().strftime("%Y-%m")
    if os.path.exists(PRODUCTS_FILE):
        pdf = pd.read_excel(PRODUCTS_FILE, engine="openpyxl")
        pdf.columns = [str(c).strip().lower() for c in pdf.columns]
        cols = ["xs","s","m","l","xl","xxl"]
        for c in cols: 
            if c in pdf.columns:
                pdf[c] = pd.to_numeric(pdf[c], errors='coerce').fillna(0).astype(int)
        pdf["total_quantity"] = pdf[[c for c in cols if c in pdf.columns]].sum(axis=1)
        for _, r in pdf.iterrows(): spends += (int(r["total_quantity"]) * float(r["purchase_price"]))
    if os.path.exists(SALES_FILE):
        sdf = pd.read_excel(SALES_FILE, engine="openpyxl")
        if not sdf.empty:
            cost_dict = {}
            if os.path.exists(PRODUCTS_FILE):
                p_df2 = pd.read_excel(PRODUCTS_FILE, engine="openpyxl")
                p_df2.columns = [str(c).strip().lower() for c in p_df2.columns]
                for _, r in p_df2.iterrows(): cost_dict[str(r["barcode"]).strip()] = float(r["purchase_price"])
            for _, r in sdf.iterrows():
                dt, rev, bc, q = str(r["date"]).strip(), float(r["total_price"]), str(r["barcode"]).strip(), int(r["quantity_sold"])
                profit += (rev - (cost_dict.get(bc, 0.0) * q))
                if dt.startswith(cur_d): t_sl += rev
                if dt.startswith(cur_m): m_sl += rev
    return render_template("reports.html", today_sales=round(t_sl,2), monthly_sales=round(m_sl,2), total_spends=round(spends,2), total_profit=round(profit,2))

@app.route("/customers")
@cond_login_required
def customers():
    try:
        sdf = pd.read_excel(SALES_FILE, engine="openpyxl") if os.path.exists(SALES_FILE) else pd.DataFrame()
        if not sdf.empty:
            sdf = sdf.dropna(subset=["customer_phone"])
            sdf["customer_phone"] = sdf["customer_phone"].astype(str).str.strip().apply(lambda val: val.split('.')[0] if val.endswith('.0') else val)
            sdf["customer_name"] = sdf["customer_name"].astype(str).str.strip()
            sdf["total_price"] = pd.to_numeric(sdf["total_price"], errors='coerce').fillna(0.0)
            sdf = sdf[sdf["customer_phone"] != ""]

            aggregated = sdf.groupby("customer_phone").agg(
                customer_name=("customer_name", "last"),
                joining_date=("date", "first"), 
                total_spend=("total_price", "sum"),
                total_visits=("transaction_id", "nunique") 
            ).reset_index()

            aggregated["customer_id"] = aggregated.index.map(lambda x: f"CUST-{x+1:03d}")
            aggregated["total_spend"] = aggregated["total_spend"].round(2)
            c = aggregated.to_dict(orient="records")
        else: c = []
    except Exception as e:
        print(f"Dynamic CRM computation engine error: {e}")
        c = []
    return render_template("customers.html", customers=c)

@app.route("/settings", methods=["GET", "POST"])
@cond_login_required
def settings():
    if request.method == "POST":
        STORE_CONFIG["owner"] = request.form.get("shop_owner", "").strip()
        STORE_CONFIG["upi"] = request.form.get("upi_id", "").strip()
        flash("Settings saved successfully, bro!")
        return redirect(url_for("settings"))
    return render_template("settings.html", config=STORE_CONFIG)

if __name__ == "__main__":
    app.run(debug=True)