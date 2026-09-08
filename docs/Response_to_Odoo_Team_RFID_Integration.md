# Response to Odoo Integration Team: RFID Panel & Hello Park Guest Registration

---

## Questions List

The integration team requested the following details for the BRD (Business Requirements Document):
1. **RFID Band Binding & Unbinding API Documentation:**
   - API endpoints and authentication details
   - Request and response specifications
   - Mandatory parameters and validation rules
   - Sample request/response payloads
   - Error codes and handling mechanism
   - Business rules and dependencies
   - Sample integration code (if available)
2. **Existing Hello Parks Guest Registration Form:**
   - The guest registration form currently used for the Hello Park setup
   - Details of all data fields captured during registration
   - Mandatory/optional fields and validation rules
   - Consent, terms & conditions, and applicable declarations
   - Customer identification and mapping logic
   - UAT credentials for the RFID Cashier Panel and Guest Registration Form

---

## PART 1. User Journey & Architecture

> 🔗 **Interactive Demonstration Environments (Demo mode, no credentials required):**  
> * 📱 **Mobile Guest Registration Form (SPA):** [https://aihrtest-create.github.io/rfid-panel/guest-registration/](https://aihrtest-create.github.io/rfid-panel/guest-registration/)  
> * 🖥️ **RFID Cashier Panel:** [https://aihrtest-create.github.io/rfid-panel/](https://aihrtest-create.github.io/rfid-panel/)  
>  
> *(These interfaces are deployed as interactive, live web prototypes. You can test the end-to-end user experience directly in your browser without login credentials).*

---

### 1.1. High-Level System Architecture (How It Works Currently)

* **Guest Registration & Game System:** Guest registration, child profiles, game progress across interactive attractions, points, and virtual Avatars are hosted and managed **on the Hello Park side**.
* **Park POS (Cash Desk):** Ticket sales and receipt printing/fiscalization take place **on the park's POS system**.
* **Controlled Wristband Issuance:** A physical RFID wristband is issued to a child **only after confirmed ticket payment at the POS**. Upon fiscalizing a receipt, the POS sends a payment signal (activation token) to Hello Park, unlocking wristband assignment. Wristbands cannot be assigned without an active paid admission ticket.

---

### 1.2. User Flow Diagram

![User Flow Diagram](./screenshots/user_flow_diagram.png)  
*(Diagram: [user_flow_diagram.png](./screenshots/user_flow_diagram.png))*

1. **Step 1 (Guest):** Upon entering the park, the guest scans a QR code using their smartphone and fills out a lightweight mobile web registration form (phone number verified via SMS OTP, parent's full name, children's names and dates of birth).
2. **Step 2 (Server):** Family profile data is automatically stored in Hello Park's global and local databases and instantly appears in the RFID Cashier Panel web interface on the cashier's computer.
3. **Step 3 (Park POS):** The guest approaches the cash desk and pays for admission tickets. At this point, the cashier cannot bind a wristband yet, as no activation tokens have been granted.
4. **Step 4 (Integration):** Immediately after receipt fiscalization/payment, the park's POS sends a lightweight HTTP POST request (webhook) to the Hello Park backend containing the receipt number and the count of admission tickets.
5. **Step 5 (Cashier):** The RFID Cashier Panel immediately receives the activation credits. The cashier selects the child, taps a clean RFID wristband against the desktop USB reader, and the wristband is bound. The child enters the park to play.

---

### 1.3. Cashier Workspace (RFID Cashier Panel)

A dedicated web application running on the cashier's computer next to the POS terminal. The interface is optimized for rapid cashier operations:

![RFID Cashier Panel Screen](./screenshots/main_screen_en.png)  
*(Screenshot: [main_screen_en.png](./screenshots/main_screen_en.png))*

* **Parent Row:** Full name, phone number, list of linked children.
* **Child Statuses:**
  * `NONE` (Grey) — Child is registered in the database, but no wristband has been assigned yet.
  * `BRACELET` (Orange) — Wristband has been linked at the desk; child is currently active in the park.
  * `AVATAR` (Purple) — Child has interacted with game projections in the park. Game progress, level, XP, and prize achievements are persistently saved to this profile.

---

### 1.4. Key Operational Scenarios

#### Scenario 1: First-Time Visit (New Family)
1. Parent fills out the online registration form upon arrival.
2. Approaches the desk; cashier processes admission tickets on the park's POS.
3. The POS triggers an payment event ➔ activation credits light up in the Hello Park panel.
4. Cashier taps a wristband on the desktop reader ➔ the wristband is linked to the child profile ➔ the child enters the attractions.

#### Scenario 2: Returning Guest (Progress Continuity)
Children frequently return to the park (next week, next month, or months later):
1. The physical wristband from their previous visit was returned upon exit, but their game progress (Avatar, level, accumulated points, unlocked rewards) **remains permanently saved in the Hello Park database**.
2. Upon return, the cashier locates the family in the system (by phone number or name).
3. The cashier takes a clean physical wristband, clicks **"New bracelet"** next to the child's profile, and taps it on the reader.
4. The new physical RFID UID is attached to the child's permanent profile (`child_id`). When the child taps the wristband at any park attraction, their **saved Avatar, level, and points are instantly restored**.

#### Scenario 3: Wristband Return & Same-Day Reuse
In high-volume parks, wristbands circulate continuously:
1. Upon exiting the park, the child drops the wristband at the desk.
2. The cashier clicks "Clear bracelet" in the panel and taps it on the reader — the wristband is unbound in 1 second. The child's game progress remains completely intact.
3. The cleared wristband is immediately returned to the pool and can be assigned to the next arriving guest.
4. **Automated Nightly Wipe:** Every night at 03:00 AM, the Hello Park server automatically resets all active physical RFID wristband associations. In the morning, all wristbands in the park are clean and ready for issuance.

---

## PART 2. Technical API Specifications (Binding & POS Integration)

### 2.1. Architectural Approaches: Who Performs Wristband Binding?

Before finalizing the BRD, we need to align on one of two implementation models:

#### Option A: Hello Park RFID Panel
* **How it works:** Cashiers use the standalone Hello Park web panel with a connected desktop USB RFID reader.
* **Role of Odoo:** Odoo **does not interact with RFID hardware directly**. Odoo only transmits a ticket payment webhook (`POST /api/v1/pos/events`).
* **Result:** Activation slots are credited to the Hello Park panel, and the cashier assigns wristbands inside the Hello Park UI.

#### Option B: Direct Binding from Odoo POS
* **How it works:** Odoo POS connects to an RFID reader directly within the Odoo POS terminal interface.
* **Role of Odoo:** After scanning the physical wristband, Odoo calls the external Hello Park API to link `child_id` + `rfid_uid`.

---

### 2.2. Specification for Option A (Payment Webhook from POS to Hello Park)

The POS sends this notification immediately upon payment/receipt fiscalization whenever the transaction includes at least one admission ticket.

* **URL:** `POST https://<HP_SERVER_URL>/api/v1/pos/events`
* **Headers:**
  ```http
  Content-Type: application/json; charset=utf-8
  Authorization: Bearer <API_SECRET_KEY>
  ```
* **Request Body (JSON):**
  ```json
  {
    "receipt_number": "REC-2026-0904-001",
    "count": 2
  }
  ```
* **Parameters:**
  * `receipt_number` *(String, Required)*: Unique receipt/transaction identifier from Odoo. Used for **idempotency** (prevents duplicate activation credits on network retries).
  * `count` *(Integer, Required)*: Total number of paid child admission tickets in the receipt.
* **Success Response (`200 OK`):**
  ```json
  {
    "status": "ok"
  }
  ```
* **Error Codes:**
  * `401 Unauthorized`: Invalid or missing secret authorization key (`API_SECRET_KEY`).
  * `400 Bad Request`: Missing mandatory parameters or malformed JSON.

---

### 2.3. Specification for Option B (Direct Bind / Unbind APIs for Odoo)

If Odoo performs the wristband scanning directly on its end:

#### Method 1: Bind Wristband
* **URL:** `POST https://<HP_SERVER_URL>/api/v1/rfid/bind`
* **Headers:** `Authorization: Bearer <API_SECRET_KEY>`, `Content-Type: application/json`
* **Request Body:**
  ```json
  {
    "child_id": "CH-10023",
    "rfid_uid": "04A1B2C3D4",
    "receipt_number": "REC-2026-0904-001"
  }
  ```
* **Success Response (`200 OK`):**
  ```json
  {
    "status": "success",
    "message": "Wristband successfully bound",
    "data": {
      "child_id": "CH-10023",
      "rfid_uid": "04A1B2C3D4",
      "avatar_status": "ready",
      "bound_at": "2026-09-04T14:30:00Z"
    }
  }
  ```
* **Error Codes:**
  * `400 Bad Request`: Missing mandatory parameters.
  * `404 Not Found`: Child with specified `child_id` not found in registry.
  * `409 Conflict`: This `rfid_uid` is already assigned to another active child in the park.
  * `422 Unprocessable Entity`: No available ticket activations remain for this receipt number.

#### Method 2: Unbind Wristband (Exit / Same-Day Reuse)
* **URL:** `POST https://<HP_SERVER_URL>/api/v1/rfid/unbind`
* **Headers:** `Authorization: Bearer <API_SECRET_KEY>`, `Content-Type: application/json`
* **Request Body:**
  ```json
  {
    "rfid_uid": "04A1B2C3D4"
  }
  ```
* **Success Response (`200 OK`):**
  ```json
  {
    "status": "success",
    "message": "Wristband unbound and released for reuse"
  }
  ```

---

### 2.4. Core Business Rules & Dependencies

1. **Separation of Physical Token and Game Profile:** The physical RFID wristband is purely a temporary day token. The child's Avatar, points, levels, and unlocked prizes are permanently tied to `child_id` in the Hello Park database and are never wiped when a wristband is removed.
2. **Returning Guests:** On subsequent visits, a new physical wristband is issued, but the binding request must pass the same permanent `child_id`.
3. **Duplicate Prevention (Conflicts):** A single wristband cannot be assigned to two children at the same time. The system returns `409 Conflict` if the wristband is already active.
4. **Automated Daily Reset:** Every night, all active physical RFID wristband assignments are cleared, leaving all hardware wristbands ready for reassignment the next morning.

---

## PART 3. Guest Registration Form (Hello Park Setup)

### 3.1. Registration Form Fields & Structure

The guest registration process is organized into sequential steps optimized for mobile browsers:

#### Step 1: Contact Information & Legal Consent
* **Phone Number** (`phone`):
  * *Requirement:* Mandatory.
  * *Format:* International phone format (mask and digit count configured according to the park's country).
  * *Verification:* 4-digit SMS OTP verification. (Simulated in demo mode; accepts any 4 digits except `0000`).
* **Terms & Legal Declarations (Checkbox):**
  * *Requirement:* Mandatory (SMS sending is disabled until accepted).
  * *Scope:* The guest confirms acceptance of:
    1. Park Visiting Rules;
    2. Loyalty Program Terms;
    3. Personal Data Processing Policy.

#### Step 2: Family & Children Details
* **Parent Full Name** (`fio`):
  * *Requirement:* Mandatory, string (First Name, Last Name, Middle Name).
* **Children List** (`children`):
  * *Requirement:* Mandatory, minimum 1 child required. Additional children can be added dynamically.
  * *Fields per child:*
    * **Child First Name** (`name`): Mandatory, text string.
    * **Date of Birth** (`dob`): Mandatory, date format (`YYYY-MM-DD` or localized date picker). Required to calibrate age-appropriate games and quests.
* **Optional / Partner-Specific Fields (Configurable per park):**
  * **"How did you hear about us?"** (Marketing attribution / acquisition channel — dropdown or text).
  * **"City"** (Guest place of residence).

#### Step 3: Data Persistence & Display (Registration Completion)
Upon form submission:
* Family data is automatically saved to Hello Park's global and local databases.
* The parent and child profiles immediately appear in the registered guest queue within the **RFID Cashier Panel**.
* The cashier looks up the family by phone number or name to assign wristbands.

> **Optional Feature (Loyalty QR Code Generation):**  
> For selected parks, an encoded digital loyalty QR card is generated on Step 3. When enabled, the cashier can simply scan this QR code with a 2D desktop scanner to open the family record instantly, avoiding manual phone searches.

---

### 3.2. Customer Identification & Mapping Logic
* **Parent Identifier (Primary Key):** Mobile phone number (`phone`).
* **Child Identifier:** Permanent `child_id` generated upon profile creation.
* **Odoo Mapping:** The Odoo customer profile should store either the external `child_id` or the primary family phone number to ensure recurring ticket purchases consistently resolve to the existing `child_id`.
