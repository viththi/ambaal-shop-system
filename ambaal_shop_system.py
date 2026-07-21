from flask import Flask, request, redirect, url_for, session, flash, render_template_string
import mysql.connector
from mysql.connector import Error
from werkzeug.security import generate_password_hash, check_password_hash
from functools import wraps
from datetime import date
from decimal import Decimal, InvalidOperation
import os

# ============================================================
# AMBAAL SHOP MANAGEMENT SYSTEM
# Everything is contained in this single Python file.
# ============================================================

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "change-this-secret-key-before-production")

# ------------------------- DATABASE --------------------------
# XAMPP/phpMyAdmin default settings are commonly:
# host = localhost, user = root, password = ""
# Change these values if your MySQL account is different.
DB_HOST = "localhost"
DB_USER = "root"
DB_PASSWORD = ""
DB_NAME = "ambaal_shop"


def server_connection():
    """Connect to MySQL without selecting a database."""
    return mysql.connector.connect(
        host=DB_HOST,
        user=DB_USER,
        password=DB_PASSWORD
    )


def db_connection():
    """Connect to the application database."""
    return mysql.connector.connect(
        host=DB_HOST,
        user=DB_USER,
        password=DB_PASSWORD,
        database=DB_NAME
    )


def initialize_database():
    """Create the database, tables and default administrator."""
    server = server_connection()
    cursor = server.cursor()
    cursor.execute(
        f"CREATE DATABASE IF NOT EXISTS `{DB_NAME}` "
        "CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
    )
    server.commit()
    cursor.close()
    server.close()

    connection = db_connection()
    cursor = connection.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INT AUTO_INCREMENT PRIMARY KEY,
            username VARCHAR(100) NOT NULL UNIQUE,
            password_hash VARCHAR(255) NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS customers (
            id INT AUTO_INCREMENT PRIMARY KEY,
            customer_code VARCHAR(50) NOT NULL UNIQUE,
            customer_name VARCHAR(150) NOT NULL,
            mobile_number VARCHAR(30) NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS loan_transactions (
            id INT AUTO_INCREMENT PRIMARY KEY,
            customer_id INT NOT NULL,
            transaction_date DATE NOT NULL,
            transaction_type ENUM('LOAN', 'PAYMENT') NOT NULL,
            amount DECIMAL(12,2) NOT NULL,
            note VARCHAR(255),
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            CONSTRAINT fk_loan_customer
                FOREIGN KEY (customer_id) REFERENCES customers(id)
                ON DELETE CASCADE
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS shop_items (
            id INT AUTO_INCREMENT PRIMARY KEY,
            item_name VARCHAR(150) NOT NULL,
            item_code VARCHAR(50) UNIQUE,
            quantity INT NOT NULL DEFAULT 0,
            description VARCHAR(255),
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS prices (
            id INT AUTO_INCREMENT PRIMARY KEY,
            item_name VARCHAR(150) NOT NULL,
            selling_price DECIMAL(12,2) NOT NULL,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                ON UPDATE CURRENT_TIMESTAMP
        )
    """)

    cursor.execute("SELECT id FROM users WHERE username = %s", ("ambaal",))
    admin = cursor.fetchone()
    if not admin:
        cursor.execute(
            "INSERT INTO users (username, password_hash) VALUES (%s, %s)",
            ("ambaal", generate_password_hash("ambaal"))
        )

    connection.commit()
    cursor.close()
    connection.close()


# ----------------------- AUTHENTICATION -----------------------

def login_required(view_function):
    @wraps(view_function)
    def wrapped_view(*args, **kwargs):
        if "user_id" not in session:
            flash("Please log in first.", "warning")
            return redirect(url_for("login"))
        return view_function(*args, **kwargs)
    return wrapped_view


# ------------------------- UTILITIES --------------------------

def money(value):
    try:
        return f"{Decimal(value):,.2f}"
    except Exception:
        return "0.00"


app.jinja_env.filters["money"] = money


def get_customer_balance(connection, customer_id):
    cursor = connection.cursor(dictionary=True)
    cursor.execute("""
        SELECT COALESCE(SUM(
            CASE
                WHEN transaction_type = 'LOAN' THEN amount
                WHEN transaction_type = 'PAYMENT' THEN -amount
                ELSE 0
            END
        ), 0) AS balance
        FROM loan_transactions
        WHERE customer_id = %s
    """, (customer_id,))
    result = cursor.fetchone()
    cursor.close()
    return Decimal(result["balance"] or 0)


# -------------------------- DESIGN ----------------------------

BASE_TEMPLATE = r"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{ title }} | Ambaal Shop</title>
    <style>
        :root {
            --primary: #5b3df5;
            --primary-dark: #4329d4;
            --sidebar: #151525;
            --background: #f4f6fb;
            --card: #ffffff;
            --text: #202235;
            --muted: #70758a;
            --success: #14804a;
            --danger: #c93838;
            --warning: #ad6800;
            --border: #e3e6ef;
        }

        * { box-sizing: border-box; }

        body {
            margin: 0;
            font-family: Arial, Helvetica, sans-serif;
            background: var(--background);
            color: var(--text);
        }

        a { text-decoration: none; color: inherit; }

        .layout {
            min-height: 100vh;
            display: flex;
        }

        .sidebar {
            width: 270px;
            background: var(--sidebar);
            color: white;
            padding: 25px 18px;
            position: fixed;
            inset: 0 auto 0 0;
            overflow-y: auto;
        }

        .brand {
            font-size: 24px;
            font-weight: 800;
            padding: 8px 12px 25px;
        }

        .brand span { color: #9c8cff; }

        .nav-label {
            margin: 18px 12px 8px;
            color: #8c8ca0;
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 1px;
        }

        .nav-link {
            display: block;
            padding: 13px 14px;
            margin: 7px 0;
            border-radius: 10px;
            color: #d8d8e3;
        }

        .nav-link:hover,
        .nav-link.active {
            background: var(--primary);
            color: white;
        }

        .main {
            margin-left: 270px;
            width: calc(100% - 270px);
            min-height: 100vh;
        }

        .topbar {
            height: 72px;
            background: white;
            border-bottom: 1px solid var(--border);
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 0 30px;
            position: sticky;
            top: 0;
            z-index: 20;
        }

        .topbar h1 {
            font-size: 21px;
            margin: 0;
        }

        .user-area {
            display: flex;
            align-items: center;
            gap: 14px;
            color: var(--muted);
        }

        .logout {
            color: var(--danger);
            font-weight: 700;
        }

        .content { padding: 28px; }

        .grid {
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 18px;
        }

        .card {
            background: var(--card);
            border: 1px solid var(--border);
            border-radius: 15px;
            padding: 22px;
            box-shadow: 0 5px 20px rgba(28, 32, 57, 0.04);
            margin-bottom: 20px;
        }

        .card h2, .card h3 { margin-top: 0; }

        .stat-title {
            color: var(--muted);
            font-size: 14px;
        }

        .stat-value {
            font-size: 30px;
            font-weight: 800;
            margin-top: 8px;
        }

        .section-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 15px;
            margin-bottom: 18px;
        }

        .section-header h2 { margin: 0; }

        .form-grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 16px;
        }

        .form-group { margin-bottom: 4px; }

        label {
            display: block;
            margin-bottom: 7px;
            font-size: 14px;
            font-weight: 700;
        }

        input, select, textarea {
            width: 100%;
            padding: 12px 13px;
            border: 1px solid #ccd1dd;
            border-radius: 9px;
            font-size: 15px;
            background: white;
        }

        input:focus, select:focus, textarea:focus {
            outline: 2px solid rgba(91, 61, 245, 0.18);
            border-color: var(--primary);
        }

        textarea { min-height: 90px; resize: vertical; }

        .full { grid-column: 1 / -1; }

        .btn {
            display: inline-block;
            border: 0;
            border-radius: 9px;
            padding: 11px 17px;
            font-size: 14px;
            font-weight: 700;
            cursor: pointer;
        }

        .btn-primary { background: var(--primary); color: white; }
        .btn-primary:hover { background: var(--primary-dark); }
        .btn-secondary { background: #eceafb; color: var(--primary-dark); }
        .btn-danger { background: #ffebeb; color: var(--danger); }
        .btn-success { background: #e4f7ed; color: var(--success); }

        .table-wrap { overflow-x: auto; }

        table {
            width: 100%;
            border-collapse: collapse;
            min-width: 760px;
        }

        th, td {
            text-align: left;
            padding: 13px 12px;
            border-bottom: 1px solid var(--border);
            font-size: 14px;
        }

        th {
            color: var(--muted);
            background: #fafbfe;
            font-size: 12px;
            text-transform: uppercase;
        }

        tr:hover td { background: #fbfbfe; }

        .badge {
            display: inline-block;
            border-radius: 999px;
            padding: 5px 9px;
            font-size: 12px;
            font-weight: 700;
        }

        .badge-loan { background: #fff0e4; color: #a94b00; }
        .badge-payment { background: #e4f7ed; color: var(--success); }
        .balance-positive { color: var(--danger); font-weight: 800; }
        .balance-zero { color: var(--success); font-weight: 800; }

        .alert {
            border-radius: 10px;
            padding: 13px 15px;
            margin-bottom: 18px;
            font-weight: 600;
        }

        .alert-success { background: #e4f7ed; color: var(--success); }
        .alert-danger { background: #ffebeb; color: var(--danger); }
        .alert-warning { background: #fff6df; color: var(--warning); }
        .alert-info { background: #e9f3ff; color: #145a96; }

        .empty {
            text-align: center;
            padding: 35px;
            color: var(--muted);
        }

        .search-row {
            display: flex;
            gap: 10px;
            margin-bottom: 18px;
        }

        .search-row input { max-width: 420px; }

        .login-page {
            min-height: 100vh;
            display: grid;
            place-items: center;
            padding: 20px;
            background:
                radial-gradient(circle at top left, #7d67ff, transparent 40%),
                linear-gradient(135deg, #151525, #292457);
        }

        .login-card {
            width: 100%;
            max-width: 430px;
            background: white;
            border-radius: 20px;
            padding: 35px;
            box-shadow: 0 22px 60px rgba(0,0,0,.25);
        }

        .login-card h1 { margin-bottom: 8px; }
        .login-card p { color: var(--muted); margin-top: 0; }
        .login-card .form-group { margin: 18px 0; }
        .login-card .btn { width: 100%; padding: 13px; }

        .mobile-menu { display: none; }

        @media (max-width: 900px) {
            .sidebar {
                transform: translateX(-100%);
                transition: .25s;
                z-index: 100;
            }

            .sidebar.open { transform: translateX(0); }

            .main {
                margin-left: 0;
                width: 100%;
            }

            .mobile-menu {
                display: inline-block;
                border: 0;
                background: #eceafb;
                color: var(--primary);
                border-radius: 8px;
                padding: 8px 11px;
                cursor: pointer;
            }

            .grid { grid-template-columns: 1fr; }
            .form-grid { grid-template-columns: 1fr; }
            .full { grid-column: auto; }
        }
    </style>
</head>
<body>
{% if logged_in %}
<div class="layout">
    <aside class="sidebar" id="sidebar">
        <div class="brand">AMBAAL <span>SHOP</span></div>

        <div class="nav-label">Main</div>
        <a class="nav-link {{ 'active' if active_page == 'dashboard' else '' }}"
           href="{{ url_for('dashboard') }}">Dashboard</a>

        <div class="nav-label">Management</div>
        <a class="nav-link {{ 'active' if active_page == 'loans' else '' }}"
           href="{{ url_for('customer_loans') }}">Customer Loan Management</a>
        <a class="nav-link {{ 'active' if active_page == 'items' else '' }}"
           href="{{ url_for('shop_items') }}">Shop Things Details</a>
        <a class="nav-link {{ 'active' if active_page == 'prices' else '' }}"
           href="{{ url_for('price_management') }}">Prices</a>
    </aside>

    <main class="main">
        <header class="topbar">
            <div style="display:flex;align-items:center;gap:12px;">
                <button class="mobile-menu" onclick="toggleMenu()">☰</button>
                <h1>{{ page_heading }}</h1>
            </div>
            <div class="user-area">
                <span>Admin: {{ session.get('username') }}</span>
                <a class="logout" href="{{ url_for('logout') }}">Logout</a>
            </div>
        </header>

        <section class="content">
            {% with messages = get_flashed_messages(with_categories=true) %}
                {% for category, message in messages %}
                    <div class="alert alert-{{ category }}">{{ message }}</div>
                {% endfor %}
            {% endwith %}

            {{ content|safe }}
        </section>
    </main>
</div>
{% else %}
<div class="login-page">
    <div class="login-card">
        <h1>Ambaal Shop</h1>
        <p>Sign in to open the administration system.</p>

        {% with messages = get_flashed_messages(with_categories=true) %}
            {% for category, message in messages %}
                <div class="alert alert-{{ category }}">{{ message }}</div>
            {% endfor %}
        {% endwith %}

        {{ content|safe }}
    </div>
</div>
{% endif %}

<script>
    function toggleMenu() {
        document.getElementById("sidebar").classList.toggle("open");
    }

    function confirmDelete(message) {
        return confirm(message || "Are you sure?");
    }
</script>
</body>
</html>
"""


def render_page(title, page_heading, content_template, active_page="", **context):
    content = render_template_string(content_template, **context)
    return render_template_string(
        BASE_TEMPLATE,
        title=title,
        page_heading=page_heading,
        content=content,
        active_page=active_page,
        logged_in=("user_id" in session)
    )


# --------------------------- LOGIN ----------------------------

@app.route("/", methods=["GET", "POST"])
def login():
    if "user_id" in session:
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")

        connection = None
        cursor = None
        try:
            connection = db_connection()
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                "SELECT id, username, password_hash FROM users WHERE username = %s",
                (username,)
            )
            user = cursor.fetchone()

            if user and check_password_hash(user["password_hash"], password):
                session.clear()
                session["user_id"] = user["id"]
                session["username"] = user["username"]
                flash("Login successful.", "success")
                return redirect(url_for("dashboard"))

            flash("Invalid username or password.", "danger")

        except Error as error:
            flash(f"Database error: {error}", "danger")
        finally:
            if cursor:
                cursor.close()
            if connection and connection.is_connected():
                connection.close()

    login_form = r"""
    <form method="POST">
        <div class="form-group">
            <label>Username</label>
            <input type="text" name="username" required autocomplete="username"
                   placeholder="Enter username">
        </div>

        <div class="form-group">
            <label>Password</label>
            <input type="password" name="password" required
                   autocomplete="current-password" placeholder="Enter password">
        </div>

        <button class="btn btn-primary" type="submit">Login</button>
    </form>
    """
    return render_page("Login", "Login", login_form)


@app.route("/logout")
def logout():
    session.clear()
    flash("You have been logged out.", "info")
    return redirect(url_for("login"))


# ------------------------- DASHBOARD --------------------------

@app.route("/dashboard")
@login_required
def dashboard():
    connection = None
    cursor = None
    try:
        connection = db_connection()
        cursor = connection.cursor(dictionary=True)

        cursor.execute("SELECT COUNT(*) AS count FROM customers")
        customer_count = cursor.fetchone()["count"]

        cursor.execute("SELECT COUNT(*) AS count FROM shop_items")
        item_count = cursor.fetchone()["count"]

        cursor.execute("""
            SELECT COALESCE(SUM(
                CASE
                    WHEN transaction_type = 'LOAN' THEN amount
                    WHEN transaction_type = 'PAYMENT' THEN -amount
                    ELSE 0
                END
            ), 0) AS total_balance
            FROM loan_transactions
        """)
        total_balance = cursor.fetchone()["total_balance"]

        cursor.execute("""
            SELECT
                lt.id,
                c.customer_name,
                c.customer_code,
                lt.transaction_date,
                lt.transaction_type,
                lt.amount
            FROM loan_transactions lt
            JOIN customers c ON c.id = lt.customer_id
            ORDER BY lt.id DESC
            LIMIT 8
        """)
        recent_transactions = cursor.fetchall()

    except Error as error:
        flash(f"Database error: {error}", "danger")
        customer_count = 0
        item_count = 0
        total_balance = 0
        recent_transactions = []
    finally:
        if cursor:
            cursor.close()
        if connection and connection.is_connected():
            connection.close()

    template = r"""
    <div class="grid">
        <div class="card">
            <div class="stat-title">Registered Customers</div>
            <div class="stat-value">{{ customer_count }}</div>
        </div>
        <div class="card">
            <div class="stat-title">Total Outstanding Loan</div>
            <div class="stat-value">Rs. {{ total_balance|money }}</div>
        </div>
        <div class="card">
            <div class="stat-title">Shop Items</div>
            <div class="stat-value">{{ item_count }}</div>
        </div>
    </div>

    <div class="card">
        <div class="section-header">
            <h2>Recent Loan Activity</h2>
            <a class="btn btn-secondary" href="{{ url_for('customer_loans') }}">
                Open Loan Management
            </a>
        </div>

        {% if recent_transactions %}
        <div class="table-wrap">
            <table>
                <thead>
                    <tr>
                        <th>Date</th>
                        <th>Customer</th>
                        <th>Customer ID</th>
                        <th>Type</th>
                        <th>Amount</th>
                    </tr>
                </thead>
                <tbody>
                {% for transaction in recent_transactions %}
                    <tr>
                        <td>{{ transaction.transaction_date }}</td>
                        <td>{{ transaction.customer_name }}</td>
                        <td>{{ transaction.customer_code }}</td>
                        <td>
                            <span class="badge {{ 'badge-loan' if transaction.transaction_type == 'LOAN' else 'badge-payment' }}">
                                {{ transaction.transaction_type }}
                            </span>
                        </td>
                        <td>Rs. {{ transaction.amount|money }}</td>
                    </tr>
                {% endfor %}
                </tbody>
            </table>
        </div>
        {% else %}
            <div class="empty">No loan transactions have been recorded.</div>
        {% endif %}
    </div>
    """

    return render_page(
        "Dashboard",
        "Dashboard",
        template,
        active_page="dashboard",
        customer_count=customer_count,
        item_count=item_count,
        total_balance=total_balance,
        recent_transactions=recent_transactions
    )


# ---------------------- CUSTOMER LOANS ------------------------

@app.route("/customer-loans")
@login_required
def customer_loans():
    search = request.args.get("search", "").strip()

    connection = None
    cursor = None
    customers = []
    transactions = []

    try:
        connection = db_connection()
        cursor = connection.cursor(dictionary=True)

        customer_query = """
            SELECT
                c.id,
                c.customer_code,
                c.customer_name,
                c.mobile_number,
                c.created_at,
                COALESCE(SUM(
                    CASE
                        WHEN lt.transaction_type = 'LOAN' THEN lt.amount
                        WHEN lt.transaction_type = 'PAYMENT' THEN -lt.amount
                        ELSE 0
                    END
                ), 0) AS balance
            FROM customers c
            LEFT JOIN loan_transactions lt ON lt.customer_id = c.id
        """
        parameters = []

        if search:
            customer_query += """
                WHERE c.customer_name LIKE %s
                   OR c.customer_code LIKE %s
                   OR c.mobile_number LIKE %s
            """
            pattern = f"%{search}%"
            parameters.extend([pattern, pattern, pattern])

        customer_query += """
            GROUP BY c.id, c.customer_code, c.customer_name,
                     c.mobile_number, c.created_at
            ORDER BY c.customer_name
        """
        cursor.execute(customer_query, tuple(parameters))
        customers = cursor.fetchall()

        cursor.execute("""
            SELECT
                lt.id,
                lt.transaction_date,
                lt.transaction_type,
                lt.amount,
                lt.note,
                lt.created_at,
                c.id AS customer_id,
                c.customer_code,
                c.customer_name,
                c.mobile_number
            FROM loan_transactions lt
            JOIN customers c ON c.id = lt.customer_id
            ORDER BY lt.transaction_date DESC, lt.id DESC
            LIMIT 100
        """)
        transactions = cursor.fetchall()

    except Error as error:
        flash(f"Database error: {error}", "danger")
    finally:
        if cursor:
            cursor.close()
        if connection and connection.is_connected():
            connection.close()

    template = r"""
    <div class="card">
        <div class="section-header">
            <h2>Customer Accounts</h2>
            <a class="btn btn-primary" href="{{ url_for('add_customer') }}">
                + Add New Customer
            </a>
        </div>

        <form class="search-row" method="GET">
            <input type="text" name="search" value="{{ search }}"
                   placeholder="Search name, customer ID or mobile number">
            <button class="btn btn-secondary" type="submit">Search</button>
            {% if search %}
                <a class="btn btn-danger" href="{{ url_for('customer_loans') }}">Clear</a>
            {% endif %}
        </form>

        {% if customers %}
        <div class="table-wrap">
            <table>
                <thead>
                    <tr>
                        <th>Customer ID</th>
                        <th>Name</th>
                        <th>Mobile</th>
                        <th>Current Balance</th>
                        <th>Actions</th>
                    </tr>
                </thead>
                <tbody>
                {% for customer in customers %}
                    <tr>
                        <td>{{ customer.customer_code }}</td>
                        <td>{{ customer.customer_name }}</td>
                        <td>{{ customer.mobile_number }}</td>
                        <td class="{{ 'balance-positive' if customer.balance > 0 else 'balance-zero' }}">
                            Rs. {{ customer.balance|money }}
                        </td>
                        <td>
                            <a class="btn btn-success"
                               href="{{ url_for('customer_details', customer_id=customer.id) }}">
                                View / Add Amount
                            </a>
                        </td>
                    </tr>
                {% endfor %}
                </tbody>
            </table>
        </div>
        {% else %}
            <div class="empty">No customers found.</div>
        {% endif %}
    </div>

    <div class="card">
        <h2>All Transaction History</h2>
        {% if transactions %}
        <div class="table-wrap">
            <table>
                <thead>
                    <tr>
                        <th>Date</th>
                        <th>Customer</th>
                        <th>Customer ID</th>
                        <th>Type</th>
                        <th>Amount</th>
                        <th>Note</th>
                    </tr>
                </thead>
                <tbody>
                {% for transaction in transactions %}
                    <tr>
                        <td>{{ transaction.transaction_date }}</td>
                        <td>{{ transaction.customer_name }}</td>
                        <td>{{ transaction.customer_code }}</td>
                        <td>
                            <span class="badge {{ 'badge-loan' if transaction.transaction_type == 'LOAN' else 'badge-payment' }}">
                                {{ transaction.transaction_type }}
                            </span>
                        </td>
                        <td>Rs. {{ transaction.amount|money }}</td>
                        <td>{{ transaction.note or '-' }}</td>
                    </tr>
                {% endfor %}
                </tbody>
            </table>
        </div>
        {% else %}
            <div class="empty">No transaction history available.</div>
        {% endif %}
    </div>
    """

    return render_page(
        "Customer Loans",
        "Customer Loan Management",
        template,
        active_page="loans",
        customers=customers,
        transactions=transactions,
        search=search
    )


@app.route("/customer-loans/add-customer", methods=["GET", "POST"])
@login_required
def add_customer():
    if request.method == "POST":
        customer_name = request.form.get("customer_name", "").strip()
        customer_code = request.form.get("customer_code", "").strip()
        mobile_number = request.form.get("mobile_number", "").strip()
        transaction_date = request.form.get("transaction_date", "")
        initial_loan = request.form.get("initial_loan", "0").strip()
        note = request.form.get("note", "").strip()

        if not customer_name or not customer_code or not mobile_number:
            flash("Customer name, customer ID and mobile number are required.", "danger")
        else:
            connection = None
            cursor = None
            try:
                amount = Decimal(initial_loan or "0")
                if amount < 0:
                    raise InvalidOperation

                connection = db_connection()
                cursor = connection.cursor()
                cursor.execute("""
                    INSERT INTO customers
                        (customer_code, customer_name, mobile_number)
                    VALUES (%s, %s, %s)
                """, (customer_code, customer_name, mobile_number))

                customer_id = cursor.lastrowid

                if amount > 0:
                    cursor.execute("""
                        INSERT INTO loan_transactions
                            (customer_id, transaction_date,
                             transaction_type, amount, note)
                        VALUES (%s, %s, 'LOAN', %s, %s)
                    """, (
                        customer_id,
                        transaction_date or date.today().isoformat(),
                        amount,
                        note or "Initial loan"
                    ))

                connection.commit()
                flash("Customer and initial loan saved successfully.", "success")
                return redirect(
                    url_for("customer_details", customer_id=customer_id)
                )

            except InvalidOperation:
                flash("Loan amount must be a valid positive number.", "danger")
            except Error as error:
                if connection:
                    connection.rollback()
                if error.errno == 1062:
                    flash("That customer ID already exists.", "danger")
                else:
                    flash(f"Database error: {error}", "danger")
            finally:
                if cursor:
                    cursor.close()
                if connection and connection.is_connected():
                    connection.close()

    template = r"""
    <div class="card">
        <div class="section-header">
            <h2>Add New Customer</h2>
            <a class="btn btn-secondary" href="{{ url_for('customer_loans') }}">Back</a>
        </div>

        <form method="POST">
            <div class="form-grid">
                <div class="form-group">
                    <label>Customer Name</label>
                    <input type="text" name="customer_name" required>
                </div>

                <div class="form-group">
                    <label>Customer ID</label>
                    <input type="text" name="customer_code" required
                           placeholder="Example: CUST-001">
                </div>

                <div class="form-group">
                    <label>Mobile Number</label>
                    <input type="tel" name="mobile_number" required>
                </div>

                <div class="form-group">
                    <label>Date</label>
                    <input type="date" name="transaction_date"
                           value="{{ today }}" required>
                </div>

                <div class="form-group">
                    <label>Initial Loan Amount (Rs.)</label>
                    <input type="number" name="initial_loan" min="0"
                           step="0.01" value="0" required>
                </div>

                <div class="form-group">
                    <label>Note</label>
                    <input type="text" name="note"
                           placeholder="Items bought or other details">
                </div>

                <div class="full">
                    <button class="btn btn-primary" type="submit">
                        Save Customer
                    </button>
                </div>
            </div>
        </form>
    </div>
    """

    return render_page(
        "Add Customer",
        "Add New Customer",
        template,
        active_page="loans",
        today=date.today().isoformat()
    )


@app.route("/customer-loans/customer/<int:customer_id>")
@login_required
def customer_details(customer_id):
    connection = None
    cursor = None

    try:
        connection = db_connection()
        cursor = connection.cursor(dictionary=True)

        cursor.execute(
            "SELECT * FROM customers WHERE id = %s",
            (customer_id,)
        )
        customer = cursor.fetchone()

        if not customer:
            flash("Customer not found.", "danger")
            return redirect(url_for("customer_loans"))

        balance = get_customer_balance(connection, customer_id)

        cursor.execute("""
            SELECT id, transaction_date, transaction_type,
                   amount, note, created_at
            FROM loan_transactions
            WHERE customer_id = %s
            ORDER BY transaction_date DESC, id DESC
        """, (customer_id,))
        transactions = cursor.fetchall()

    except Error as error:
        flash(f"Database error: {error}", "danger")
        return redirect(url_for("customer_loans"))
    finally:
        if cursor:
            cursor.close()
        if connection and connection.is_connected():
            connection.close()

    template = r"""
    <div class="grid">
        <div class="card">
            <div class="stat-title">Customer</div>
            <div class="stat-value" style="font-size:23px;">
                {{ customer.customer_name }}
            </div>
            <p>{{ customer.customer_code }} · {{ customer.mobile_number }}</p>
        </div>

        <div class="card">
            <div class="stat-title">Current Outstanding Balance</div>
            <div class="stat-value {{ 'balance-positive' if balance > 0 else 'balance-zero' }}">
                Rs. {{ balance|money }}
            </div>
        </div>

        <div class="card">
            <div class="stat-title">Number of Transactions</div>
            <div class="stat-value">{{ transactions|length }}</div>
        </div>
    </div>

    <div class="card">
        <div class="section-header">
            <h2>Add Loan or Payment</h2>
            <a class="btn btn-secondary" href="{{ url_for('customer_loans') }}">
                Back to Customers
            </a>
        </div>

        <form method="POST"
              action="{{ url_for('add_transaction', customer_id=customer.id) }}">
            <div class="form-grid">
                <div class="form-group">
                    <label>Transaction Type</label>
                    <select name="transaction_type" required>
                        <option value="LOAN">New Loan / Credit Purchase</option>
                        <option value="PAYMENT">Customer Payment</option>
                    </select>
                </div>

                <div class="form-group">
                    <label>Date</label>
                    <input type="date" name="transaction_date"
                           value="{{ today }}" required>
                </div>

                <div class="form-group">
                    <label>Amount (Rs.)</label>
                    <input type="number" name="amount" min="0.01"
                           step="0.01" required>
                </div>

                <div class="form-group">
                    <label>Note</label>
                    <input type="text" name="note"
                           placeholder="Example: Rice and groceries">
                </div>

                <div class="full">
                    <button class="btn btn-primary" type="submit">
                        Save Transaction
                    </button>
                </div>
            </div>
        </form>
    </div>

    <div class="card">
        <h2>{{ customer.customer_name }} - History</h2>

        {% if transactions %}
        <div class="table-wrap">
            <table>
                <thead>
                    <tr>
                        <th>Date</th>
                        <th>Type</th>
                        <th>Loan Added</th>
                        <th>Payment</th>
                        <th>Note</th>
                        <th>Recorded At</th>
                        <th>Action</th>
                    </tr>
                </thead>
                <tbody>
                {% for transaction in transactions %}
                    <tr>
                        <td>{{ transaction.transaction_date }}</td>
                        <td>
                            <span class="badge {{ 'badge-loan' if transaction.transaction_type == 'LOAN' else 'badge-payment' }}">
                                {{ transaction.transaction_type }}
                            </span>
                        </td>
                        <td>
                            {% if transaction.transaction_type == 'LOAN' %}
                                Rs. {{ transaction.amount|money }}
                            {% else %}-{% endif %}
                        </td>
                        <td>
                            {% if transaction.transaction_type == 'PAYMENT' %}
                                Rs. {{ transaction.amount|money }}
                            {% else %}-{% endif %}
                        </td>
                        <td>{{ transaction.note or '-' }}</td>
                        <td>{{ transaction.created_at }}</td>
                        <td>
                            <form method="POST"
                                  action="{{ url_for('delete_transaction', transaction_id=transaction.id) }}"
                                  onsubmit="return confirmDelete('Delete this transaction?');">
                                <button class="btn btn-danger" type="submit">Delete</button>
                            </form>
                        </td>
                    </tr>
                {% endfor %}
                </tbody>
            </table>
        </div>
        {% else %}
            <div class="empty">No transactions for this customer.</div>
        {% endif %}
    </div>
    """

    return render_page(
        "Customer Details",
        customer["customer_name"],
        template,
        active_page="loans",
        customer=customer,
        balance=balance,
        transactions=transactions,
        today=date.today().isoformat()
    )


@app.route(
    "/customer-loans/customer/<int:customer_id>/transaction",
    methods=["POST"]
)
@login_required
def add_transaction(customer_id):
    transaction_type = request.form.get("transaction_type", "").upper()
    transaction_date = request.form.get("transaction_date", "")
    amount_text = request.form.get("amount", "").strip()
    note = request.form.get("note", "").strip()

    if transaction_type not in ("LOAN", "PAYMENT"):
        flash("Invalid transaction type.", "danger")
        return redirect(url_for("customer_details", customer_id=customer_id))

    try:
        amount = Decimal(amount_text)
        if amount <= 0:
            raise InvalidOperation
    except (InvalidOperation, ValueError):
        flash("Amount must be greater than zero.", "danger")
        return redirect(url_for("customer_details", customer_id=customer_id))

    connection = None
    cursor = None
    try:
        connection = db_connection()

        if transaction_type == "PAYMENT":
            current_balance = get_customer_balance(connection, customer_id)
            if amount > current_balance:
                flash(
                    f"Payment cannot be greater than the current balance "
                    f"(Rs. {money(current_balance)}).",
                    "danger"
                )
                return redirect(
                    url_for("customer_details", customer_id=customer_id)
                )

        cursor = connection.cursor()
        cursor.execute("""
            INSERT INTO loan_transactions
                (customer_id, transaction_date,
                 transaction_type, amount, note)
            VALUES (%s, %s, %s, %s, %s)
        """, (
            customer_id,
            transaction_date or date.today().isoformat(),
            transaction_type,
            amount,
            note or None
        ))
        connection.commit()

        new_balance = get_customer_balance(connection, customer_id)
        flash(
            f"Transaction saved. Current balance: Rs. {money(new_balance)}",
            "success"
        )

    except Error as error:
        if connection:
            connection.rollback()
        flash(f"Database error: {error}", "danger")
    finally:
        if cursor:
            cursor.close()
        if connection and connection.is_connected():
            connection.close()

    return redirect(url_for("customer_details", customer_id=customer_id))


@app.route(
    "/customer-loans/transaction/<int:transaction_id>/delete",
    methods=["POST"]
)
@login_required
def delete_transaction(transaction_id):
    connection = None
    cursor = None
    customer_id = None

    try:
        connection = db_connection()
        cursor = connection.cursor(dictionary=True)
        cursor.execute(
            "SELECT customer_id FROM loan_transactions WHERE id = %s",
            (transaction_id,)
        )
        transaction = cursor.fetchone()

        if not transaction:
            flash("Transaction not found.", "danger")
            return redirect(url_for("customer_loans"))

        customer_id = transaction["customer_id"]
        cursor.execute(
            "DELETE FROM loan_transactions WHERE id = %s",
            (transaction_id,)
        )
        connection.commit()
        flash("Transaction deleted.", "success")

    except Error as error:
        if connection:
            connection.rollback()
        flash(f"Database error: {error}", "danger")
    finally:
        if cursor:
            cursor.close()
        if connection and connection.is_connected():
            connection.close()

    if customer_id:
        return redirect(url_for("customer_details", customer_id=customer_id))
    return redirect(url_for("customer_loans"))


# ------------------------- SHOP ITEMS -------------------------

@app.route("/shop-items", methods=["GET", "POST"])
@login_required
def shop_items():
    if request.method == "POST":
        item_name = request.form.get("item_name", "").strip()
        item_code = request.form.get("item_code", "").strip() or None
        quantity = request.form.get("quantity", "0").strip()
        description = request.form.get("description", "").strip()

        connection = None
        cursor = None
        try:
            quantity_value = int(quantity)
            if quantity_value < 0:
                raise ValueError

            if not item_name:
                flash("Item name is required.", "danger")
            else:
                connection = db_connection()
                cursor = connection.cursor()
                cursor.execute("""
                    INSERT INTO shop_items
                        (item_name, item_code, quantity, description)
                    VALUES (%s, %s, %s, %s)
                """, (
                    item_name,
                    item_code,
                    quantity_value,
                    description or None
                ))
                connection.commit()
                flash("Shop item saved successfully.", "success")
                return redirect(url_for("shop_items"))

        except ValueError:
            flash("Quantity must be zero or a positive whole number.", "danger")
        except Error as error:
            if connection:
                connection.rollback()
            if error.errno == 1062:
                flash("That item code already exists.", "danger")
            else:
                flash(f"Database error: {error}", "danger")
        finally:
            if cursor:
                cursor.close()
            if connection and connection.is_connected():
                connection.close()

    connection = None
    cursor = None
    items = []
    try:
        connection = db_connection()
        cursor = connection.cursor(dictionary=True)
        cursor.execute("SELECT * FROM shop_items ORDER BY id DESC")
        items = cursor.fetchall()
    except Error as error:
        flash(f"Database error: {error}", "danger")
    finally:
        if cursor:
            cursor.close()
        if connection and connection.is_connected():
            connection.close()

    template = r"""
    <div class="card">
        <h2>Add Shop Thing</h2>
        <form method="POST">
            <div class="form-grid">
                <div class="form-group">
                    <label>Item Name</label>
                    <input type="text" name="item_name" required>
                </div>

                <div class="form-group">
                    <label>Item Code</label>
                    <input type="text" name="item_code"
                           placeholder="Optional unique code">
                </div>

                <div class="form-group">
                    <label>Quantity</label>
                    <input type="number" name="quantity" min="0" value="0" required>
                </div>

                <div class="form-group">
                    <label>Description</label>
                    <input type="text" name="description">
                </div>

                <div class="full">
                    <button class="btn btn-primary" type="submit">Save Item</button>
                </div>
            </div>
        </form>
    </div>

    <div class="card">
        <h2>Shop Things History</h2>
        {% if items %}
        <div class="table-wrap">
            <table>
                <thead>
                    <tr>
                        <th>Item Code</th>
                        <th>Item Name</th>
                        <th>Quantity</th>
                        <th>Description</th>
                        <th>Created</th>
                        <th>Action</th>
                    </tr>
                </thead>
                <tbody>
                {% for item in items %}
                    <tr>
                        <td>{{ item.item_code or '-' }}</td>
                        <td>{{ item.item_name }}</td>
                        <td>{{ item.quantity }}</td>
                        <td>{{ item.description or '-' }}</td>
                        <td>{{ item.created_at }}</td>
                        <td>
                            <form method="POST"
                                  action="{{ url_for('delete_item', item_id=item.id) }}"
                                  onsubmit="return confirmDelete('Delete this item?');">
                                <button class="btn btn-danger" type="submit">Delete</button>
                            </form>
                        </td>
                    </tr>
                {% endfor %}
                </tbody>
            </table>
        </div>
        {% else %}
            <div class="empty">No shop things have been added.</div>
        {% endif %}
    </div>
    """

    return render_page(
        "Shop Things",
        "Shop Things Details",
        template,
        active_page="items",
        items=items
    )


@app.route("/shop-items/<int:item_id>/delete", methods=["POST"])
@login_required
def delete_item(item_id):
    connection = None
    cursor = None
    try:
        connection = db_connection()
        cursor = connection.cursor()
        cursor.execute("DELETE FROM shop_items WHERE id = %s", (item_id,))
        connection.commit()
        flash("Shop item deleted.", "success")
    except Error as error:
        if connection:
            connection.rollback()
        flash(f"Database error: {error}", "danger")
    finally:
        if cursor:
            cursor.close()
        if connection and connection.is_connected():
            connection.close()

    return redirect(url_for("shop_items"))


# ---------------------- PRICE MANAGEMENT ----------------------

@app.route("/prices", methods=["GET", "POST"])
@login_required
def price_management():
    if request.method == "POST":
        item_name = request.form.get("item_name", "").strip()
        price_text = request.form.get("selling_price", "").strip()

        connection = None
        cursor = None
        try:
            selling_price = Decimal(price_text)
            if selling_price < 0:
                raise InvalidOperation

            if not item_name:
                flash("Item name is required.", "danger")
            else:
                connection = db_connection()
                cursor = connection.cursor()
                cursor.execute("""
                    INSERT INTO prices (item_name, selling_price)
                    VALUES (%s, %s)
                """, (item_name, selling_price))
                connection.commit()
                flash("Price saved successfully.", "success")
                return redirect(url_for("price_management"))

        except (InvalidOperation, ValueError):
            flash("Selling price must be a valid positive number.", "danger")
        except Error as error:
            if connection:
                connection.rollback()
            flash(f"Database error: {error}", "danger")
        finally:
            if cursor:
                cursor.close()
            if connection and connection.is_connected():
                connection.close()

    connection = None
    cursor = None
    prices = []
    try:
        connection = db_connection()
        cursor = connection.cursor(dictionary=True)
        cursor.execute("SELECT * FROM prices ORDER BY id DESC")
        prices = cursor.fetchall()
    except Error as error:
        flash(f"Database error: {error}", "danger")
    finally:
        if cursor:
            cursor.close()
        if connection and connection.is_connected():
            connection.close()

    template = r"""
    <div class="card">
        <h2>Add Item Price</h2>
        <form method="POST">
            <div class="form-grid">
                <div class="form-group">
                    <label>Item Name</label>
                    <input type="text" name="item_name" required>
                </div>

                <div class="form-group">
                    <label>Selling Price (Rs.)</label>
                    <input type="number" name="selling_price"
                           min="0" step="0.01" required>
                </div>

                <div class="full">
                    <button class="btn btn-primary" type="submit">Save Price</button>
                </div>
            </div>
        </form>
    </div>

    <div class="card">
        <h2>Price List</h2>
        {% if prices %}
        <div class="table-wrap">
            <table>
                <thead>
                    <tr>
                        <th>Item Name</th>
                        <th>Selling Price</th>
                        <th>Last Updated</th>
                        <th>Action</th>
                    </tr>
                </thead>
                <tbody>
                {% for price in prices %}
                    <tr>
                        <td>{{ price.item_name }}</td>
                        <td>Rs. {{ price.selling_price|money }}</td>
                        <td>{{ price.updated_at }}</td>
                        <td>
                            <form method="POST"
                                  action="{{ url_for('delete_price', price_id=price.id) }}"
                                  onsubmit="return confirmDelete('Delete this price?');">
                                <button class="btn btn-danger" type="submit">Delete</button>
                            </form>
                        </td>
                    </tr>
                {% endfor %}
                </tbody>
            </table>
        </div>
        {% else %}
            <div class="empty">No prices have been added.</div>
        {% endif %}
    </div>
    """

    return render_page(
        "Prices",
        "Price Management",
        template,
        active_page="prices",
        prices=prices
    )


@app.route("/prices/<int:price_id>/delete", methods=["POST"])
@login_required
def delete_price(price_id):
    connection = None
    cursor = None
    try:
        connection = db_connection()
        cursor = connection.cursor()
        cursor.execute("DELETE FROM prices WHERE id = %s", (price_id,))
        connection.commit()
        flash("Price deleted.", "success")
    except Error as error:
        if connection:
            connection.rollback()
        flash(f"Database error: {error}", "danger")
    finally:
        if cursor:
            cursor.close()
        if connection and connection.is_connected():
            connection.close()

    return redirect(url_for("price_management"))


# --------------------------- START ----------------------------

if __name__ == "__main__":
    try:
        initialize_database()
        print("=" * 55)
        print("AMBAAL SHOP MANAGEMENT SYSTEM")
        print("Open: http://0.0.0.0:5003")
        print("Username: ambaal")
        print("Password: ambaal")
        print("=" * 55)
        app.run(debug=False, host="0.0.0.0", port=5003)
    except Error as error:
        print("\nCould not connect to MySQL.")
        print("Please start Apache and MySQL in XAMPP.")
        print("Also check DB_HOST, DB_USER and DB_PASSWORD in app.py.")
        print(f"Error details: {error}")