import random
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd


# -----------------------------
# Configuration
# -----------------------------

random.seed(42)

BASE_DIR = Path(r"D:\Projects\powerbi-sales-sample")
OUTPUT_DIR = BASE_DIR / "Data"
OUTPUT_FILE = OUTPUT_DIR / "sales_data.xlsx"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

NUM_LEADS = 2000

salespeople = [
    "John Smith",
    "Sarah Johnson",
    "Michael Brown",
    "Emily Davis",
    "David Wilson",
]

lead_sources = [
    "Google",
    "Facebook",
    "Referral",
    "Website",
    "HomeAdvisor",
    "Yard Sign",
    "Repeat Customer",
    "Other",
]

project_types = [
    "Kitchen Remodel",
    "Bathroom Remodel",
    "Roofing",
    "Flooring",
    "Home Addition",
    "Exterior Renovation",
]

first_names = [
    "James", "Robert", "John", "Michael", "David",
    "William", "Richard", "Joseph", "Thomas", "Charles",
    "Mary", "Patricia", "Jennifer", "Linda", "Elizabeth",
    "Barbara", "Susan", "Jessica", "Sarah", "Karen",
]

last_names = [
    "Anderson", "Brown", "Clark", "Davis", "Evans",
    "Garcia", "Harris", "Johnson", "Jones", "Martin",
    "Miller", "Moore", "Robinson", "Smith", "Taylor",
    "Thomas", "Thompson", "Walker", "White", "Wilson",
]

# Approximate contract value ranges by project type
project_value_ranges = {
    "Kitchen Remodel": (12000, 35000),
    "Bathroom Remodel": (7000, 22000),
    "Roofing": (9000, 28000),
    "Flooring": (5000, 18000),
    "Home Addition": (30000, 90000),
    "Exterior Renovation": (10000, 40000),
}

# Lead source quality affects likelihood of winning
source_win_probability = {
    "Google": 0.16,
    "Facebook": 0.09,
    "Referral": 0.25,
    "Website": 0.14,
    "HomeAdvisor": 0.11,
    "Yard Sign": 0.08,
    "Repeat Customer": 0.30,
    "Other": 0.07,
}

# Salesperson performance affects close rate
salesperson_multiplier = {
    "John Smith": 1.05,
    "Sarah Johnson": 1.20,
    "Michael Brown": 0.88,
    "Emily Davis": 1.12,
    "David Wilson": 0.96,
}


# -----------------------------
# Generate data
# -----------------------------

start_date = datetime(2025, 8, 1)
end_date = datetime(2026, 7, 31)

date_range_days = (end_date - start_date).days

rows = []

for i in range(1, NUM_LEADS + 1):

    lead_date = start_date + timedelta(
        days=random.randint(0, date_range_days)
    )

    lead_source = random.choices(
        lead_sources,
        weights=[22, 16, 18, 14, 10, 7, 8, 5],
        k=1,
    )[0]

    salesperson = random.choice(salespeople)

    project_type = random.choice(project_types)

    customer = (
        f"{random.choice(first_names)} "
        f"{random.choice(last_names)}"
    )

    min_value, max_value = project_value_ranges[project_type]

    contract_value = round(
        random.uniform(min_value, max_value),
        -2
    )

    # Determine status
    base_probability = source_win_probability[lead_source]
    adjusted_probability = (
        base_probability
        * salesperson_multiplier[salesperson]
    )

    # Some leads remain open
    status_roll = random.random()

    if status_roll < 0.18:
        status = "Open"
    elif random.random() < adjusted_probability:
        status = "Won"
    else:
        status = "Lost"

    # Sales process dates
    appointment_date = None
    proposal_date = None
    close_date = None
    sales_cycle_days = None

    if status in ["Won", "Lost"]:

        appointment_days = random.randint(1, 7)

        appointment_date = lead_date + timedelta(
            days=appointment_days
        )

        proposal_date = appointment_date + timedelta(
            days=random.randint(2, 14)
        )

        close_date = proposal_date + timedelta(
            days=random.randint(3, 35)
        )

        sales_cycle_days = (
            close_date.date() - lead_date.date()
        ).days

    elif status == "Open":

        # Some open leads have appointments/proposals
        if random.random() < 0.65:

            appointment_date = lead_date + timedelta(
                days=random.randint(1, 7)
            )

        if appointment_date and random.random() < 0.55:

            proposal_date = appointment_date + timedelta(
                days=random.randint(2, 14)
            )

    rows.append(
        {
            "Lead ID": f"L{i:05d}",
            "Lead Date": lead_date.date(),
            "Lead Source": lead_source,
            "Salesperson": salesperson,
            "Customer": customer,
            "Project Type": project_type,
            "Contract Value": contract_value,
            "Status": status,
            "Appointment Date": (
                appointment_date.date()
                if appointment_date
                else None
            ),
            "Proposal Date": (
                proposal_date.date()
                if proposal_date
                else None
            ),
            "Close Date": (
                close_date.date()
                if close_date
                else None
            ),
            "Sales Cycle Days": sales_cycle_days,
        }
    )


# -----------------------------
# Create Excel file
# -----------------------------

df = pd.DataFrame(rows)

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl",
    date_format="yyyy-mm-dd",
) as writer:

    df.to_excel(
        writer,
        sheet_name="Leads",
        index=False,
    )


print()
print("======================================")
print("Sales dataset created successfully!")
print("======================================")
print()
print(f"File: {OUTPUT_FILE}")
print()
print(f"Rows: {len(df):,}")
print()
print("Status distribution:")
print(df["Status"].value_counts())
print()
print("Lead sources:")
print(df["Lead Source"].value_counts())
print()