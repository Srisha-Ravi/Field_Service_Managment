from flask import Flask, jsonify, render_template, request
import mysql.connector
from datetime import datetime
from flask_cors import CORS
import os
from fpdf import FPDF
import win32com.client as win32


app = Flask(__name__)
CORS(app)  # For JS fetch requests

# Global variable to hold connection
db_connection = None

def get_db_connection():
    global db_connection
    
    # Close existing connection if it exists
    if db_connection is not None and db_connection.is_connected():
        db_connection.close()
        print("Existing connection closed.")
    
    # Open a new connection
    db_connection = mysql.connector.connect(
        host="localhost",
        user="root",
        password="Harsh@1997",
        database="ltservice"
    )
    print("New connection opened.")
    return db_connection


# ------------------- API -------------------
# ------------------- Create new customer -------------------
@app.route("/customers", methods=["POST"])
def create_customer():
    data = request.json
    print("Incoming JSON:", data)  # Debug log
    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        # Insert customer
        cursor.execute("""
            INSERT INTO tblcustomermaster (customername, customershortname)
            VALUES (%s, %s)
        """, (data.get("customername"), data.get("customershortname")))
        customer_id = cursor.lastrowid

        # Insert sites and machines
        for site in data.get("sites", []):
            cursor.execute("""
                INSERT INTO tblsitemaster
                (customerid, sitename, site_shortname, addr1, addr2, city, state, pincode, phone, email, gstno)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            """, (
                customer_id,
                site.get("sitename"),
                site.get("site_shortname"),
                site.get("addr1"),
                site.get("addr2"),
                site.get("city"),
                site.get("state"),
                site.get("pincode"),
                site.get("phone"),
                site.get("email"),
                site.get("gstno")
            ))
            site_id = cursor.lastrowid

            for m in site.get("machines", []):
                cursor.execute("""
                    INSERT INTO tblmachinemaster
                    (siteid, customerid, machineno, machinetype, make, model, amc_expiry_date)
                    VALUES (%s,%s,%s,%s,%s,%s,%s)
                """, (
                    site_id, 
                    customer_id,
                    m.get("machineno"),
                    m.get("machinetype"),
                    m.get("make"),
                    m.get("model"),
                    m.get("amc_expiry_date")
                ))

        conn.commit()
        return jsonify({"status": "success", "customerid": customer_id})

    except Exception as e:
        conn.rollback()
        return str(e), 500

    finally:
        cursor.close()
        conn.close()
# ------------------- Update existing customer -------------------
@app.route("/customers/<int:customer_id>", methods=["PUT"])
def update_customer(customer_id):
    data = request.json
    print("Incoming JSON:", data)
    print("📌 customername:", data.get("customername"))
    print("📌 customershortname:", data.get("customershortname"))
    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        # ---- Update Customer Info ----
        cursor.execute("""
            UPDATE tblcustomermaster
            SET customername=%s, customershortname=%s
            WHERE customerid=%s
        """, (data.get("customername"), data.get("customershortname"), customer_id))

        # ---- Handle Sites ----
        incoming_sites = data.get("sites", [])

        # Fetch existing sites from DB
        cursor.execute("SELECT siteid, site_shortname FROM tblsitemaster WHERE customerid=%s", (customer_id,))
        db_sites = {row[1]: row[0] for row in cursor.fetchall()}  # {site_shortname: siteid}

        incoming_shortnames = [s.get("site_shortname") for s in incoming_sites]

        # Delete sites that are NOT in incoming JSON
        for shortname, siteid in db_sites.items():
            if shortname not in incoming_shortnames:
                cursor.execute("DELETE FROM tblmachinemaster WHERE siteid=%s", (siteid,))
                cursor.execute("DELETE FROM tblsitemaster WHERE siteid=%s", (siteid,))

        # Insert/Update sites
        for site in incoming_sites:
            site_shortname = site.get("site_shortname")
            if site_shortname in db_sites:
                # Update existing site
                siteid = db_sites[site_shortname]
                cursor.execute("""
                    UPDATE tblsitemaster
                    SET sitename=%s, addr1=%s, addr2=%s, city=%s, state=%s, pincode=%s, phone=%s, email=%s, gstno=%s
                    WHERE siteid=%s
                """, (
                    site.get("sitename"),
                    site.get("addr1"),
                    site.get("addr2"),
                    site.get("city"),
                    site.get("state"),
                    site.get("pincode"),
                    site.get("phone"),
                    site.get("email"),
                    site.get("gstno"),
                    siteid
                ))
            else:
                # Insert new site
                cursor.execute("""
                    INSERT INTO tblsitemaster
                    (customerid, sitename, site_shortname, addr1, addr2, city, state, pincode, phone, email, gstno)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                """, (
                    customer_id,
                    site.get("sitename"),
                    site_shortname,
                    site.get("addr1"),
                    site.get("addr2"),
                    site.get("city"),
                    site.get("state"),
                    site.get("pincode"),
                    site.get("phone"),
                    site.get("email"),
                    site.get("gstno")
                ))
                siteid = cursor.lastrowid

            # ---- Handle Machines for this site ----
            incoming_machines = site.get("machines", [])
            cursor.execute("SELECT machineid, machineno FROM tblmachinemaster WHERE siteid=%s", (siteid,))
            db_machines = {row[1]: row[0] for row in cursor.fetchall()}  # {machineno: machineid}

            incoming_nos = [m.get("machineno") for m in incoming_machines]

            # Delete old machines not in incoming
            for machineno, mid in db_machines.items():
                if machineno not in incoming_nos:
                    cursor.execute("DELETE FROM tblmachinemaster WHERE machineid=%s", (mid,))

            # Insert/Update machines
            for m in incoming_machines:
                machineno = m.get("machineno")
                if machineno in db_machines:
                    # Update
                    mid = db_machines[machineno]
                    cursor.execute("""
                        UPDATE tblmachinemaster
                        SET machinetype=%s, make=%s, model=%s, amc_expiry_date=%s
                        WHERE machineid=%s
                    """, (
                        m.get("machinetype"),
                        m.get("make"),
                        m.get("model"),
                        m.get("amc_expiry_date"),
                        mid
                    ))
                else:
                    # Insert new
                    cursor.execute("""
                        INSERT INTO tblmachinemaster
                        (siteid, customerid, machineno, machinetype, make, model, amc_expiry_date)
                        VALUES (%s,%s,%s,%s,%s,%s,%s)
                    """, (
                        siteid,
                        customer_id,
                        machineno,
                        m.get("machinetype"),
                        m.get("make"),
                        m.get("model"),
                        m.get("amc_expiry_date")
                    ))

        conn.commit()
        return jsonify({"status": "success"})

    except Exception as e:
        conn.rollback()
        import traceback
        print(traceback.format_exc())
        return str(e), 500

    finally:
        cursor.close()
        conn.close()

# ------------------- Get customer for editing -------------------
@app.route("/customers/<int:customer_id>", methods=["GET"])
def get_customer(customer_id):
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    try:
        cursor.execute("SELECT * FROM tblcustomermaster WHERE customerid=%s", (customer_id,))
        customer = cursor.fetchone()
        if not customer:
            return "Customer not found", 404

        cursor.execute("SELECT * FROM tblsitemaster WHERE customerid=%s", (customer_id,))
        sites = cursor.fetchall()
        for site in sites:
            cursor.execute("SELECT * FROM tblmachinemaster WHERE siteid=%s", (site["siteid"],))
            site["machines"] = cursor.fetchall()
        customer["sites"] = sites
        return jsonify(customer)

    finally:
        cursor.close()
        conn.close()

 #--------to get all customer list-----------

@app.route("/customers", methods=["GET"])
def list_customers():
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT * FROM tblcustomermaster")
    customers = cursor.fetchall()
    print(customers)
    for customer in customers:
        cursor.execute("SELECT * FROM tblsitemaster WHERE customerid=%s", (customer["customerid"],))
        sites = cursor.fetchall()
        for site in sites:
            cursor.execute("SELECT * FROM tblmachinemaster WHERE siteid=%s", (site["siteid"],))
            site["machines"] = cursor.fetchall()
        customer["sites"] = sites
    cursor.close()
    conn.close()
    
    return jsonify(customers)
       

# ------------------- Delete customer -------------------

@app.route("/customers/<int:id>", methods=["DELETE"])
def delete_customer(id):
    conn = get_db_connection()
    cursor = conn.cursor()

    # Delete machines
    cursor.execute("SELECT siteid FROM tblsitemaster WHERE customerid=%s", (id,))
    site_ids = [row[0] for row in cursor.fetchall()]
    if site_ids:
    #cursor.execute("DELETE FROM tblmachinemaster WHERE siteid IN (%s)" % ','.join(map(str, site_ids)))
        cursor.executemany("DELETE FROM tblmachinemaster WHERE siteid=%s", [(sid,) for sid in site_ids])


    # Delete sites
    cursor.execute("DELETE FROM tblsitemaster WHERE customerid=%s", (id,))

    # Delete customer
    cursor.execute("DELETE FROM tblcustomermaster WHERE customerid=%s", (id,))

    conn.commit()
    conn.close()
    return jsonify({"message": "Customer deleted successfully"})

# ------------------- In active customer -------------------

@app.route("/customers/<int:id>/inactive", methods=["PUT"])
def inactive_customer(id):
    conn = get_db_connection()
    cursor = conn.cursor()

    # Inactive machines
    cursor.execute("SELECT siteid FROM tblsitemaster WHERE customerid=%s", (id,))
    site_ids = [row[0] for row in cursor.fetchall()]
    if site_ids:
        cursor.executemany("UPDATE tblmachinemaster SET isactive='N' WHERE siteid=%s", [(sid,) for sid in site_ids])


    # Inactive sites
    cursor.execute("UPDATE tblsitemaster SET isactive='N' WHERE customerid=%s", (id,))

    # Inactive customer
    cursor.execute("UPDATE tblcustomermaster SET isactive='N' WHERE customerid=%s", (id,))

    conn.commit()
    conn.close()
    return jsonify({"message": "Customer inactivated successfully"})


# ------------------- Complaints API -------------------

# List all complaints
@app.route("/api/complaints", methods=["GET"])
def list_complaints():
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    query = """
    SELECT 
        c.complaintid,
        c.complaintdate, 
        cm.customername,
        cm.customerid,
        s.sitename,
        s.siteid,
        m.machineno,
        m.machineid,
        c.customercomplaint,
        c.status
    FROM tblcomplaint c
    JOIN tblcustomermaster cm ON c.customerid = cm.customerid
    JOIN tblsitemaster s ON c.siteid = s.siteid
    JOIN tblmachinemaster m ON c.machineid = m.machineid
    """
    cursor.execute(query)
    results = cursor.fetchall()
    print(results)

# ✅ Convert complaintdate to DD-MM-YYYY format
    complaints = []
    for r in results:
        r["complaintdate"] = r["complaintdate"].strftime("%Y-%m-%d")  # same as example
        complaints.append(r)

    cursor.close()
    conn.close()
    print(complaints)
    return jsonify(complaints)


# Create new complaint
@app.route("/api/complaints", methods=["POST"])
def create_complaint():
    data = request.json
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO tblcomplaint 
        (complaintdate, customerid, siteid, machineid, customercomplaint, status, contactperson, contactno, inspection, repairdone) 
        VALUES (%s, %s, %s, %s, %s, 'Open', %s, %s, %s, %s)
    """, (
        datetime.strptime(data.get("complaintdate"), "%Y-%m-%d").date(),       
        data.get("customerid"),
        data.get("siteid"),
        data.get("machineid"),
        data.get("customercomplaint"),
        data.get("contactperson"),
        data.get("contactno"),
        data.get("inspection"),
        data.get("repairdone")
    ))

    complaint_id = cursor.lastrowid
    conn.commit()
    cursor.close()
    conn.close()
    return jsonify({"message": "Complaint saved", "complaintid": complaint_id})


# Update complaint
@app.route("/api/complaints/<int:id>", methods=["PUT"])
def update_complaint(id):
    data = request.json
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE tblcomplaint
        SET complaintdate=%s,customerid=%s, siteid=%s, machineid=%s, customercomplaint=%s,
            contactperson=%s, contactno=%s, inspection=%s, repairdone=%s, status=%s
        WHERE complaintid=%s
    """, (
        datetime.strptime(data.get("complaintdate"), "%Y-%m-%d").date(),
        data.get("customerid"),
        data.get("siteid"),
        data.get("machineid"),
        data.get("customercomplaint"),
        data.get("contactperson"),
        data.get("contactno"),
        data.get("inspection"),
        data.get("repairdone"),
        data.get("status"),
        id
    ))
    conn.commit()
    cursor.close()
    conn.close()
    return jsonify({"message": "Complaint updated","complaintid": id})


# Delete complaint
@app.route("/api/complaints/<int:id>", methods=["DELETE"])
def delete_complaint(id):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM tblcomplaint WHERE complaintid=%s", (id,))
    conn.commit()
    cursor.close()
    conn.close()
    return jsonify({"message": "Complaint deleted"})


# ------------------- Complaint Closure -------------------
@app.route("/api/complaints/<int:id>/close", methods=["POST"])
def close_complaint(id):
    data = request.json
    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        # ✅ Update complaint with closure info
        cursor.execute("""
            UPDATE tblcomplaint
            SET 
                inspection=%s,
                repairdone=%s,
                closuremode=%s,
                closedate=%s,
                contactperson=%s,
                contactno=%s,
                instructor=%s,
                ac=%s,
                ups=%s,
                mainpower=%s,
                cleanliness=%s,
                status='Closed'
            WHERE complaintid=%s
        """, (
            data.get("inspection"),
            data.get("repairdone"),
            data.get("closuremode"),
            datetime.now().date(),  # close date = today
            data.get("contactperson"),
            data.get("contactno"),
            data.get("instructor", "Y"),   # default Y if not sent
            data.get("ac", "Y"),
            data.get("ups", "Y"),
            data.get("mainpower", "Y"),
            data.get("cleanliness", "Y"),
            id
        ))

        # ✅ Insert replaced parts
        for part in data.get("parts", []):
            cursor.execute("""
                INSERT INTO tblpartsreplaced (complaintid, component, partno, partname, quantity)
                VALUES (%s, %s, %s, %s, %s)
            """, (
                id,
                part.get("component"),
                part.get("partno"),
                part.get("partname"),
                part.get("quantity")
            ))

        conn.commit()
        return jsonify({"status": "success", "message": "Complaint closed"})

    except Exception as e:
        conn.rollback()
        return jsonify({"status": "error", "message": str(e)}), 500

    finally:
        cursor.close()
        conn.close()



# Complaint closure page ,new

#prefilling
@app.route('/api/complaints/<int:id>', methods=['GET'])
def get_complaint(id):
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    try:
        cursor.execute("""
            SELECT 
                c.complaintid, 
                c.complaintdate,  
                cu.customername,
                s.sitename, 
                s.addr1,
                s.addr2,
                s.city,
                s.state,
                s.pincode,
                s.phone,
                s.email, 
                m.machineno, 
                m.make, 
                m.model, 
                m.machinetype,
                c.customercomplaint, 
                c.status,
                c.engineername, 
                c.inspection, 
                c.repairdone,
                c.closuremode, 
                c.closedate,
                c.contactperson, 
                c.contactno,
                c.instructor, 
                c.ac, 
                c.ups, 
                c.mainpower, 
                c.cleanliness
            FROM tblcomplaint c
            LEFT JOIN tblcustomermaster cu ON c.customerid = cu.customerid
            LEFT JOIN tblsitemaster s ON c.siteid = s.siteid
            LEFT JOIN tblmachinemaster m ON c.machineid = m.machineid
            WHERE c.complaintid = %s
        """, (id,))
        complaint = cursor.fetchone()

        if not complaint:
            return jsonify({"error": "Complaint not found"}), 404
        
        # ✅ Fetch replaced parts
        cursor.execute("""
            SELECT component, partno, partname, quantity
            FROM tblpartsreplaced
            WHERE complaintid = %s
        """, (id,))
        parts = cursor.fetchall()
        complaint["parts"] = parts   # attach to JSON response

        # ✅ Format dates only if they are real dates
        if complaint.get("complaintdate") and hasattr(complaint["complaintdate"], "strftime"):
            complaint["complaintdate"] = complaint["complaintdate"].strftime("%Y-%m-%d")
        if complaint.get("closedate") and hasattr(complaint["closedate"], "strftime"):
            complaint["closedate"] = complaint["closedate"].strftime("%Y-%m-%d")

        print("Complaint fetched:", complaint)

        return jsonify(complaint)

    except Exception as e:
        import traceback
        print("❌ Error fetching complaint:", e)
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

    finally:
        cursor.close()
        conn.close()

# ------------------- Dropdown APIs -------------------

# Get all companies
@app.route("/companies", methods=["GET"])
def get_companies():
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT customerid, customername FROM tblcustomermaster WHERE isactive='Y'")
    companies = cursor.fetchall()
    cursor.close()
    conn.close()
    return jsonify(companies)


# Get sites for a company
@app.route("/companies/<int:customer_id>/sites", methods=["GET"])
def get_sites(customer_id):
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT siteid, sitename FROM tblsitemaster WHERE customerid=%s AND isactive='Y'", (customer_id,))
    sites = cursor.fetchall()
    cursor.close()
    conn.close()
    return jsonify(sites)


# Get machines for a site
@app.route("/sites/<int:site_id>/machines", methods=["GET"])
def get_machines(site_id):
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT machineid, machineno FROM tblmachinemaster WHERE siteid=%s AND isactive='Y'", (site_id,))
    machines = cursor.fetchall()
    cursor.close()
    conn.close()
    return jsonify(machines)



# ------------------- FRONTEND ROUTES -------------------

@app.route("/")
def home():
    return render_template("main.html")

@app.route("/customer_list")
def customer_list():
    return render_template("customer_list.html")

@app.route("/form")
def form():
    return render_template("customer_form.html")

@app.route("/complaints_form")
def service_complaint():
    return render_template("service_complaint.html")

# ✅ Add POST endpoint for saving complaints
@app.route("/save_complaint", methods=["POST"])
def save_complaint():
    customer_name = request.form.get("customer_name")
    complaint_text = request.form.get("complaint_text")

    # TODO: save to DB (for now just simulate success)
    if customer_name and complaint_text:
        return jsonify({"status": "success", "message": "Complaint saved successfully"})
    else:
        return jsonify({"status": "error", "message": "Missing data"}), 400
    

@app.route('/send_email', methods=['POST'])
def send_email():
    data = request.get_json()
    cid = data.get("id")

    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT complaintid, complaintdate FROM tblcomplaint WHERE complaintid=%s", (cid,))
        result = cursor.fetchone()
        cursor.close()
        conn.close()

        if not result:
            return jsonify({"message": "Complaint not found"})

        complaint_id, complaint_date = result

        # --- Create PDF ---
        pdf_path = os.path.join(
            os.environ["USERPROFILE"], "Documents", f"complaint_{complaint_id}.pdf"
        )

        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 10, txt="Complaint Details", ln=True, align="C")
        pdf.ln(10)
        pdf.cell(200, 10, txt=f"Complaint ID: {complaint_id}", ln=True)
        pdf.cell(200, 10, txt=f"Complaint Date: {complaint_date}", ln=True)
        pdf.output(pdf_path)

        # --- Open Outlook Draft ---
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # MailItem
        mail.To = "receiver@example.com"  # could also come from DB or frontend
        mail.Subject = f"Lewov Tech Complaint {complaint_id}"
        mail.Body = (
            "Dear Sir,\n\n"
            "Please find attached the service report.\n\n"
            "Regards,\nLewov Tech"
        )
        mail.Attachments.Add(pdf_path)
        mail.Display()  # opens Outlook draft

        return jsonify({"message": "Outlook draft opened with PDF attached"})

    except Exception as e:
        return jsonify({"message": f"Error: {str(e)}"})

if __name__ == "__main__":
    app.run(debug=True)
