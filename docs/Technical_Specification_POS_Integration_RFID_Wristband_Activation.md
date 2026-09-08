# TECHNICAL SPECIFICATION
## POS Terminal Integration with HP Avatar Server

---

## 1. System Architecture and User Journey

### Key Concept:
Guest registration and user database management are hosted **on our side**. After registration, the system waits for a ticket sale event (activation token) from the POS terminal to enable RFID wristband binding to a child profile in our game system (to preserve game progress across interactive park attractions).

### User Flow:

![User Flow](screenshots/user_flow_diagram.png)

---

### Step 1. Guest fills out registration form (on our side)

Upon entering the park, the guest scans a **QR code** with their smartphone and accesses a mobile web registration form. In the form, the guest:
* Enters phone number and verifies it via SMS.
* Agrees to park rules and loyalty program terms.
* Enters parent's full name and adds children details (name, date of birth).

---

### Step 2. Data is stored in our database and appears in RFID Panel

Upon form submission, family data is automatically saved to **our database** and instantly appears in the web interface of the **RFID Cashier Panel**.

**The RFID Cashier Panel** is a web application running on the cashier's computer at the park. It displays a list of registered guest accounts:
* **Parent Account:** Full name, phone number.
* **Linked Children:** Name, date of birth, wristband status (Avatar / Bracelet / None).

![RFID Cashier Panel](screenshots/main_screen_en.png)

---

### Step 3. Guest pays for ticket at POS

The guest approaches the ticket desk and pays for admission tickets. At this stage, the cashier **cannot** bind a wristband to a child because no activation token is available yet.

**The activation token** is an authorization grant for wristband binding. It appears in the RFID panel **only after the POS terminal confirms the payment** by sending a request to our backend. Without an activation token, the wristband binding action is locked. This ensures wristbands are only issued to paying guests.

---

### Step 4. POS terminal sends request to our backend (Integration)

This is the only step requiring implementation on the POS software side. Immediately after receipt fiscalization/printing, the POS terminal sends a **single simple HTTP POST request** to our server (see specification in Section 2.1 below).

---

### Step 5. Cashier binds wristband to child

Upon receiving the event from the POS terminal, the **activation token appears instantly** in the RFID panel (available wristband counter increments). The cashier:
1. Selects the child from the list.
2. Taps a clean RFID wristband against the desktop USB/RFID reader.
3. The wristband is linked to the child's profile — the child can now enter and play.

---

## 2. API Request Specifications

### 2.1 Receiving Payment Data from POS Terminal

The POS vendor must implement sending a ticket payment signal to our local server.

* **When to send:** Immediately upon successful payment and receipt printing/fiscalization on the POS terminal.
* **Trigger condition:** When the receipt contains at least 1 admission ticket (item from the 'Tickets' category). If the receipt contains only bar/café items or merchandise, no request is sent.

#### Request:

* **Method:** `POST`
* **URL:** `https://<HP_LOCAL_SERVER_URL>/api/v1/pos/events`
* **Headers:**
  ```http
  Content-Type: application/json; charset=utf-8
  Authorization: Bearer <API_SECRET_KEY>
  ```
  *(where `<API_SECRET_KEY>` is a secret authorization token for the POS issued by us upon connection. Without a valid key, the server rejects the request with HTTP 401 Unauthorized)*

#### JSON Payload:

```json
{
  "receipt_number": "12345",
  "count": 2
}
```

#### Parameters Description:

| Parameter | Type | Required | Description & Example |
| :--- | :--- | :---: | :--- |
| `receipt_number` | String | **Yes** | Printed receipt number (e.g., `"12345"` or `"00123"`) |
| `count` | Integer | **Yes** | Total number of paid tickets in the receipt (e.g., `1`, `2`, `3`) |

#### Server Response Format:

Upon successful processing, the backend responds with **`200 OK`**:

```json
{
  "status": "ok"
}
```

If the authorization token is missing or invalid, the server responds with **`401 Unauthorized`** and rejects the operation.

#### Fault Tolerance & Retries:

* **Idempotency (Duplicate Protection):** Transactions are uniquely identified by `receipt_number`. If a network timeout causes the POS terminal to retry sending the same receipt, the system will not create duplicate activation credits and will safely return `200 OK`.
* **Offline Queue on POS:** If the backend is temporarily unreachable when closing a receipt, the POS module must store the payload in a local queue and automatically retry sending once connectivity is restored.

---

### 2.2 Guest Lookup & Authorization at POS

If the POS terminal needs to retrieve guest data, it sends a GET request to our cloud server:

```http
GET https://<HP_SERVER>/api/user
```

**Request parameters (query or payload):**
* **User ID** (scanned from the QR code shown to the guest upon registration)
* **OR Phone number**

In response, our cloud server returns the guest profile details and linked children.

---

### 2.3 Turnstiles and Access Control System (ACS / Time Tracking)

After wristband assignment, our system can send a notification request to a third-party Access Control System (ACS / Turnstiles):

**POST request parameters:**
* `user_id` — User / Child identifier
* `rfid` — RFID wristband number
* `date` — Registration timestamp / date

---

### 2.4 New Guest Registration Webhook (CRM Integration)

Upon registration of a new guest, our server can trigger a webhook on a third-party CRM system to transfer user profile details:

* `user_id` — ID in our system
* `contacts` — Phone number / contact details
* `data` — Profile data (full name, gender, age, date of birth)
