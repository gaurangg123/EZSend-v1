# 📧 Outlook VBA Macro – Send Fresh Drafts with Random Delay

This VBA macro automates sending emails from the **Outbox** in Microsoft Outlook by:  
- Cleaning recipient addresses (To, CC, BCC)  
- Preserving subject & HTML body formatting  
- Assigning a **randomized delay (1–5 minutes)** between each send  
- Ensuring all emails are sent **before 11:30 PM** (configurable)  
- Deleting the original Outbox emails after scheduling (to avoid duplicates)  

---

## ✨ Features
- ✅ Random delay between each email (1–5 minutes)  
- ✅ Deadline cutoff (default: 11:30 PM) – avoids scheduling past the day  
- ✅ Cleans invalid characters from recipient fields  
- ✅ Sends emails in HTML format (no formatting issues)  
- ✅ Safe exit messages when Outbox is empty or cutoff is reached  

---

## 📂 Macro Code
The main macro is called:  
```vb
SendFreshDraftsWithRandomDelay_CleanedHTML
```

Utility function included:  
```vb
CleanEmail()
```

---

## 🛠️ Setup Instructions

### Step 1: Open Outlook VBA Editor
1. Open **Microsoft Outlook**.  
2. Press **`Alt + F11`** to open the VBA editor.  
3. In the left pane, expand **Project1 (VbaProject.OTM)**.  

---

### Step 2: Insert the Macro
1. Go to **Insert > Module**.  
2. Copy–paste the macro code into the new module.  
3. Save the project (`Ctrl + S`).  

---

### Step 3: Add a Quick Access Button (Optional)
1. In Outlook, right-click the ribbon → **Customize the Ribbon**.  
2. Create a new group under **Home** (e.g., “Macros”).  
3. Add the macro `SendFreshDraftsWithRandomDelay_CleanedHTML` to this group.  
4. (Optional) Assign an icon for easy access.  

---

## ▶️ How to Run
- Ensure you have emails in **Outbox** (drafts waiting to send).  
- Run the macro:
  - From VBA Editor: Press **F5**  
  - From Outlook Ribbon: Click your assigned button  
- The macro will:
  - Pick each draft  
  - Schedule with a randomized delay (up to 5 mins each)  
  - Stop once 11:30 PM cutoff is reached  
  - Delete originals after scheduling  

---

## ⚠️ Notes
- You can change the cutoff time by modifying:  
  ```vb
  deadline = Date + TimeValue("23:30:00")
  ```
- Default delay range = 1–5 minutes. You can adjust inside:  
  ```vb
  randomDelay = Int((5 - 1 + 1) * Rnd + 1)
  ```
- Original drafts are **deleted after processing** (to avoid duplicates). If you want to keep them, remove this line:  
  ```vb
  originalMail.Delete
  ```

---

## 🖼️ Demo Workflow
1. Put your drafts into **Outbox**  
2. Run the macro  
3. Watch as each email gets a randomized delivery time before 11:30 PM  
4. Outbox cleans up automatically  

---

## ✅ Example Success Message
When all drafts are processed, you’ll see:  
> **"All fresh drafts sent with random delays before 11:30 PM."**  

---

## 📌 Requirements
- Microsoft Outlook (Desktop, Windows)  
- Macros enabled (Trust Center → Macro Settings)  
- Basic familiarity with VBA  
