import pandas as pd
import random
from faker import Faker
from datetime import date

fake = Faker()

kpis = [
    "Monthly Sales", "Customer Service", "Leads Generated", "Average Response Time",
    "NPS", "Conversion Rate", "Average Ticket", "Churn Rate", "Team Productivity",
    "Retention Rate", "Ticket Volume", "Resolution Time", "Bugs Fixed",
    "Internal Satisfaction", "Training Completed", "Projects Delivered", "Cost Per Acquisition",
    "Social Media Engagement", "Number of Meetings", "Positive Feedback"
]

responsibles = ["John", "Michael", "Sarah", "Jessica", "David", "Emily", "James", "Emma", "Robert", "Olivia"]

categories = ["Sales", "Marketing", "Support", "Finance", "Development", "HR", "Logistics"]

num_entries = 500

data = {
    "KPI": [random.choice(kpis) for _ in range(num_entries)],
    "Responsible": [random.choice(responsibles) for _ in range(num_entries)],
    "Target": [],
    "Current Result": [],
    "Deadline": [fake.date_between(start_date="today", end_date="+90d") for _ in range(num_entries)],
    "Category": [random.choice(categories) for _ in range(num_entries)],
}

def generate_target_result(kpi):
    if kpi in ["Monthly Sales", "Leads Generated", "Number of Meetings", "Bugs Fixed", "Projects Delivered", "Ticket Volume", "Training Completed", "Positive Feedback"]:
        target = random.randint(50, 500)
        result = max(0, int(random.gauss(target * 0.8, target * 0.2)))
    elif kpi in ["Average Response Time", "Resolution Time"]:
        target = round(random.uniform(1, 5), 2)  # hours, lower is better
        result = max(0.5, round(random.gauss(target, 1), 2))
    elif kpi in ["NPS", "Internal Satisfaction", "Customer Service"]:
        target = random.randint(80, 100)
        result = max(0, min(100, int(random.gauss(target, 10))))
    elif kpi in ["Conversion Rate", "Retention Rate", "Social Media Engagement"]:
        target = round(random.uniform(5, 20), 2)  # percent
        result = max(0, min(100, round(random.gauss(target, 5), 2)))
    elif kpi == "Average Ticket":
        target = round(random.uniform(100, 1000), 2)
        result = max(0, round(random.gauss(target, target * 0.3), 2))
    elif kpi == "Churn Rate":
        target = round(random.uniform(1, 10), 2)  # percent, lower is better
        result = max(0, round(random.gauss(target, 3), 2))
    else:
        target = random.randint(10, 100)
        result = max(0, int(random.gauss(target, 10)))
    return target, result

def determine_status(kpi, target, result, deadline):
    today = date.today()
    days_left = (deadline - today).days

    lower_is_better = kpi in ["Average Response Time", "Resolution Time", "Churn Rate"]

    if lower_is_better:
        if result <= target:
            return "Achieved"
        elif days_left > 7:
            return "In Progress"
        elif 0 <= days_left <= 7:
            return "At Risk"
        else:
            return "Delayed"
    else:
        if result >= target:
            return "Achieved"
        elif days_left > 7:
            return "In Progress"
        elif 0 <= days_left <= 7:
            return "At Risk"
        else:
            return "Delayed"

statuses = []

for i in range(num_entries):
    t, r = generate_target_result(data["KPI"][i])
    data["Target"].append(t)
    data["Current Result"].append(r)
    status = determine_status(data["KPI"][i], t, r, data["Deadline"][i])
    statuses.append(status)

data["Status"] = statuses

df_kpis = pd.DataFrame(data)

df_kpis.to_excel("team_kpis_mock.xlsx", index=False)

print("File 'team_kpis_mock.xlsx' created!")