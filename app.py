from flask import Flask, render_template, request, redirect, url_for, flash, session
from openpyxl import Workbook
from database import get_connection
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from flask import send_file

import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = "employee_management_secret"

UPLOAD_FOLDER = "static/uploads"

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER


# LOGIN

@app.route("/login", methods=["GET", "POST"])
def login():

    if request.method == "POST":

        username = request.form["username"]
        password = request.form["password"]

        conn = get_connection()
        cursor = conn.cursor(dictionary=True)

        cursor.execute(
            """
            SELECT *
            FROM admin
            WHERE username=%s
            AND password=%s
            """,
            (username, password)
        )

        admin = cursor.fetchone()

        cursor.close()
        conn.close()

        if admin:

            session["admin"] = username

            return redirect(url_for("dashboard"))

        flash("Invalid Username or Password", "danger")

    return render_template("login.html")


# CHANGE PASSWORD

@app.route("/change_password", methods=["GET", "POST"])
def change_password():

    if "admin" not in session:
        return redirect(url_for("login"))

    if request.method == "POST":

        old_password = request.form["old_password"]
        new_password = request.form["new_password"]
        confirm_password = request.form["confirm_password"]

        conn = get_connection()
        cursor = conn.cursor(dictionary=True)

        cursor.execute(
            """
            SELECT *
            FROM admin
            WHERE username=%s
            AND password=%s
            """,
            (session["admin"], old_password)
        )

        admin = cursor.fetchone()

        if not admin:

            flash("Old Password is incorrect!", "danger")

        elif new_password != confirm_password:

            flash("New passwords do not match!", "warning")

        else:

            cursor.execute(
                """
                UPDATE admin
                SET password=%s
                WHERE username=%s
                """,
                (new_password, session["admin"])
            )

            conn.commit()

            flash("Password Changed Successfully!", "success")

        cursor.close()
        conn.close()

    return render_template("change_password.html")


# HOME PAGE

@app.route("/")
def home():
    return redirect(url_for("login"))


# DASHBOARD

@app.route("/dashboard")
def dashboard():

    if "admin" not in session:
        return redirect(url_for("login"))

    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    cursor.execute("SELECT COUNT(*) AS total FROM employees")
    total_employees = cursor.fetchone()["total"]

    cursor.execute("SELECT COUNT(DISTINCT department) AS total FROM employees")
    total_departments = cursor.fetchone()["total"]

    cursor.execute("SELECT IFNULL(SUM(salary),0) AS total FROM employees")
    total_salary = cursor.fetchone()["total"]

    cursor.execute("SELECT IFNULL(AVG(salary),0) AS average_salary FROM employees")
    average_salary = round(cursor.fetchone()["average_salary"], 2)

    cursor.execute("""
        SELECT *
        FROM employees
        ORDER BY id DESC
        LIMIT 5
    """)

    latest_employees = cursor.fetchall()

    cursor.execute("""
        SELECT department, COUNT(*) AS total
        FROM employees
        GROUP BY department
        """)

    department_data = cursor.fetchall()

    cursor.execute("""
    SELECT name, salary
    FROM employees
    ORDER BY salary DESC
    LIMIT 1
    """)

    highest_employee = cursor.fetchone()

    cursor.execute("""
    SELECT COUNT(*) AS total
    FROM employees
    WHERE MONTH(joining_date)=MONTH(CURDATE())
    AND YEAR(joining_date)=YEAR(CURDATE())
    """)

    new_employees = cursor.fetchone()["total"]

    cursor.close()
    conn.close()

    return render_template(
        "dashboard.html",
        total_employees=total_employees,
        total_departments=total_departments,
        total_salary=total_salary,
        average_salary=average_salary,
        latest_employees=latest_employees,
        department_data=department_data,
        highest_employee=highest_employee,
        new_employees=new_employees
        )


# VIEW EMPLOYEES

@app.route("/employees")
def employees():

    if "admin" not in session:
        return redirect(url_for("login"))

    page = request.args.get("page", 1, type=int)
    per_page = 5

    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    cursor.execute("SELECT COUNT(*) AS total FROM employees")
    total_records = cursor.fetchone()["total"]

    total_pages = (total_records + per_page - 1) // per_page

    offset = (page - 1) * per_page

    cursor.execute("""
        SELECT *
        FROM employees
        ORDER BY id DESC
        LIMIT %s OFFSET %s
    """, (per_page, offset))

    employee_list = cursor.fetchall()

    cursor.execute("""
        SELECT DISTINCT department
        FROM employees
        ORDER BY department
    """)

    departments = cursor.fetchall()

    cursor.close()
    conn.close()

    return render_template(
        "employees.html",
        employees=employee_list,
        departments=departments,
        page=page,
        total_pages=total_pages
    )



# SEARCH EMPLOYEE

@app.route("/search")
def search():

    if "admin" not in session:
        return redirect(url_for("login"))

    keyword = request.args.get("keyword", "")
    department = request.args.get("department", "")
    page = request.args.get("page", 1, type=int)

    per_page = 5
    offset = (page - 1) * per_page

    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    value = "%" + keyword + "%"

    count_sql = """
    SELECT COUNT(*) AS total
    FROM employees
    WHERE
    (name LIKE %s
     OR email LIKE %s
     OR department LIKE %s)
    """

    count_values = [value, value, value]

    if department:
        count_sql += " AND department=%s"
        count_values.append(department)

    cursor.execute(count_sql, tuple(count_values))
    total_records = cursor.fetchone()["total"]

    total_pages = (total_records + per_page - 1) // per_page

    sql = """
    SELECT *
    FROM employees
    WHERE
    (name LIKE %s
     OR email LIKE %s
     OR department LIKE %s)
    """

    values = [value, value, value]

    if department:
        sql += " AND department=%s"
        values.append(department)

    sql += " ORDER BY id DESC LIMIT %s OFFSET %s"

    values.extend([per_page, offset])

    cursor.execute(sql, tuple(values))
    employee_list = cursor.fetchall()

    cursor.execute("""
        SELECT DISTINCT department
        FROM employees
        ORDER BY department
    """)

    departments = cursor.fetchall()

    cursor.close()
    conn.close()

    return render_template(
        "employees.html",
        employees=employee_list,
        keyword=keyword,
        departments=departments,
        selected_department=department,
        page=page,
        total_pages=total_pages
    )

# ADD EMPLOYEE

@app.route("/add_employee", methods=["GET", "POST"])
def add_employee():

    if "admin" not in session:
        return redirect(url_for("login"))

    if request.method == "POST":

        name = request.form["name"]
        email = request.form["email"]
        phone = request.form["phone"]
        department = request.form["department"]
        salary = request.form["salary"]
        joining_date = request.form["joining_date"]

        photo = request.files["photo"]

        filename = ""

        if photo.filename != "":

            filename = secure_filename(photo.filename)

            photo.save(
                os.path.join(
                    app.config["UPLOAD_FOLDER"],
                    filename
                )
            )

        conn = get_connection()
        cursor = conn.cursor()

        sql = """
        INSERT INTO employees
        (name,email,phone,department,salary,joining_date,photo)
        VALUES(%s,%s,%s,%s,%s,%s,%s)
        """

        values = (
            name,
            email,
            phone,
            department,
            salary,
            joining_date,
            filename
        )

        cursor.execute(sql, values)

        conn.commit()

        cursor.close()
        conn.close()

        flash("Employee Added Successfully!", "success")

        return redirect(url_for("employees"))

    return render_template("add_employee.html")



# EDIT EMPLOYEE

@app.route("/edit_employee/<int:id>", methods=["GET", "POST"])
def edit_employee(id):

    if "admin" not in session:
        return redirect(url_for("login"))

    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    if request.method == "POST":

        name = request.form["name"]
        email = request.form["email"]
        phone = request.form["phone"]
        department = request.form["department"]
        salary = request.form["salary"]
        joining_date = request.form["joining_date"]

        sql = """
        UPDATE employees
        SET
            name=%s,
            email=%s,
            phone=%s,
            department=%s,
            salary=%s,
            joining_date=%s
        WHERE id=%s
        """

        cursor.execute(
            sql,
            (
                name,
                email,
                phone,
                department,
                salary,
                joining_date,
                id
            )
        )

        conn.commit()

        flash("Employee Updated Successfully!", "success")

        cursor.close()
        conn.close()

        return redirect(url_for("employees"))

    cursor.execute(
        "SELECT * FROM employees WHERE id=%s",
        (id,)
    )

    employee = cursor.fetchone()

    cursor.close()
    conn.close()

    return render_template(
        "edit_employee.html",
        employee=employee
    )



# DELETE EMPLOYEE

@app.route("/delete_employee/<int:id>")
def delete_employee(id):

    if "admin" not in session:
        return redirect(url_for("login"))

    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        "DELETE FROM employees WHERE id=%s",
        (id,)
    )

    conn.commit()

    cursor.close()
    conn.close()

    flash("Employee Deleted Successfully!", "warning")

    return redirect(url_for("employees"))

@app.route("/test")
def test():
    return "Route Working"


# ERROR PAGE

@app.errorhandler(404)
def page_not_found(error):
    return render_template("404.html"), 404


# LOGOUT

@app.route("/logout")
def logout():

    session.pop("admin", None)

    flash("Logged Out Successfully!", "success")

    return redirect(url_for("login"))


# VIEW EMPLOYEE PROFILE

@app.route("/employee/<int:id>")
def employee_profile(id):

    if "admin" not in session:
        return redirect(url_for("login"))

    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    cursor.execute(
        """
        SELECT *
        FROM employees
        WHERE id=%s
        """,
        (id,)
    )

    employee = cursor.fetchone()

    cursor.close()
    conn.close()

    return render_template(
        "employee_profile.html",
        employee=employee
    )

@app.route("/export_excel")
def export_excel():

    if "admin" not in session:
        return redirect(url_for("login"))

    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    cursor.execute("""
        SELECT
            id,
            name,
            email,
            phone,
            department,
            salary,
            joining_date
        FROM employees
        ORDER BY id
    """)

    employees = cursor.fetchall()

    cursor.close()
    conn.close()

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Employees"

    sheet.append([
        "ID",
        "Name",
        "Email",
        "Phone",
        "Department",
        "Salary",
        "Joining Date"
    ])

    for employee in employees:

        sheet.append([
            employee["id"],
            employee["name"],
            employee["email"],
            employee["phone"],
            employee["department"],
            employee["salary"],
            str(employee["joining_date"])
        ])

    filename = "employees.xlsx"

    workbook.save(filename)

    return send_file(
        filename,
        as_attachment=True
    )

@app.route("/export_pdf")
def export_pdf():

    if "admin" not in session:
        return redirect(url_for("login"))

    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    cursor.execute("""
        SELECT
            id,
            name,
            email,
            phone,
            department,
            salary,
            joining_date
        FROM employees
        ORDER BY id
    """)

    employees = cursor.fetchall()

    cursor.close()
    conn.close()

    filename = "employees.pdf"

    pdf = SimpleDocTemplate(filename)

    data = [[
        "ID",
        "Name",
        "Email",
        "Phone",
        "Department",
        "Salary",
        "Joining Date"
    ]]

    for employee in employees:

        data.append([
            employee["id"],
            employee["name"],
            employee["email"],
            employee["phone"],
            employee["department"],
            str(employee["salary"]),
            str(employee["joining_date"])
        ])

    table = Table(data)

    table.setStyle(TableStyle([

        ('BACKGROUND', (0,0), (-1,0), colors.darkblue),

        ('TEXTCOLOR', (0,0), (-1,0), colors.white),

        ('GRID', (0,0), (-1,-1), 1, colors.black),

        ('BACKGROUND', (0,1), (-1,-1), colors.beige),

        ('ALIGN', (0,0), (-1,-1), 'CENTER'),

        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),

        ('BOTTOMPADDING', (0,0), (-1,0), 10)

    ]))

    pdf.build([table])

    return send_file(
        filename,
        as_attachment=True
    )


# APPLICATION ENTRY POINT

if __name__ == "__main__":

    app.run(
        host="127.0.0.1",
        port=5000,
        debug=True
    )