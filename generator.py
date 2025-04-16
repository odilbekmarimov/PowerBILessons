import pandas as pd
import numpy as np
from faker import Faker
import random
from datetime import datetime, timedelta
from pandas import ExcelWriter

# Initialize Faker for generating realistic fake data
fake = Faker()

# Number of samples to generate
num_samples = 1000

# Lists of possible values based on the table definitions
production_lines = ['Line1', 'Line2', 'Line3', 'Line4', 'Line5']
shifts = ['Day', 'Swing', 'Night']
categories = ['Electronic', 'Mechanical', 'Consumer']
defect_types = ['Assembly', 'Material', 'Finish']
severity_levels = ['Low', 'Medium', 'High']
customer_types = ['Distributors', 'Retailers', 'Direct']
countries = ['USA', 'Canada', 'Germany', 'UK']
statuses = ['Running', 'Idle', 'Setup', 'Maintenance']
units_of_measure = ['kg', 'liter', 'piece']
maintenance_types = ['Preventive', 'Corrective', 'Predictive']
transaction_types = ['Revenue', 'Material Cost', 'Labor', 'Overhead']
shipment_statuses = ['In Transit', 'Delivered', 'Delayed']
energy_types = ['Electricity', 'Gas', 'Water']
energy_units = ['kWh', 'cubic meters', 'liters']

# Helper functions
def random_date(start_year, end_year):
    start_date = datetime(start_year, 1, 1)
    end_date = datetime(end_year, 12, 31)
    delta = end_date - start_date
    random_days = random.randint(0, delta.days)
    return (start_date + timedelta(days=random_days)).strftime('%Y-%m-%d')

def random_date_str():
    return fake.date_between(start_date='-2y', end_date=datetime.today()).strftime('%Y-%m-%d')

def random_datetime_str():
    return fake.date_time_between(start_date='-2y', end_date=datetime.today()).strftime('%Y-%m-%d %H:%M:%S')

def random_month_str():
    return fake.date_between(start_date='-2y', end_date=datetime.today()).strftime('%Y-%m')

def random_time_str():
    return fake.time(pattern='%H:%M')

# Shared IDs for linking tables
product_ids = [f'P00{i % 100 + 1}' for i in range(num_samples)]  # P001-P100
customer_ids = [f'C00{i % 100 + 1}' for i in range(num_samples)]  # C001-C100
machine_ids = [f'M00{i % 50 + 1}' for i in range(num_samples)]  # M001-M050
supplier_ids = [f'S00{i % 50 + 1}' for i in range(num_samples)]  # S001-S050
material_ids = [f'MAT{i % 50 + 1}' for i in range(num_samples)]  # MAT1-MAT50
department_ids = [f'DEPT{i % 10 + 1}' for i in range(num_samples)]  # DEPT1-DEPT10
shift_ids = [f'SHFT{i % 3 + 1}' for i in range(num_samples)]  # SHFT1-SHFT3
invoice_ids = [f'INV-{1000 + i}' for i in range(num_samples)]  # INV-1000 to INV-1999

# Project 1: Products (new table)
products_data = {
    'ProductID': product_ids,
    'ProductName': [fake.word().capitalize() + ' ' + str(i % 100 + 1) for i in range(num_samples)],
    'Category': [random.choice(categories) for _ in range(num_samples)],
    'ProductCost(USD)': [round(random.uniform(10, 500), 2) for _ in range(num_samples)],
}
products_df = pd.DataFrame(products_data)

# Project 1: Production (updated to remove product metadata)
production_data = {
    'Date': [random_date_str() for _ in range(num_samples)],
    'ProductionLine': [random.choice(production_lines) for _ in range(num_samples)],
    'Shift': [random.choice(shifts) for _ in range(num_samples)],
    'ProductID': product_ids,
    'UnitsProduced': [random.randint(50, 500) for _ in range(num_samples)],
    'Defects': [random.randint(0, 20) for _ in range(num_samples)],
    'PlannedProduction': [random.randint(100, 600) for _ in range(num_samples)],
    'OperationalHours': [round(random.uniform(1, 24), 1) for _ in range(num_samples)],
    'Downtime(min)': [random.randint(0, 120) for _ in range(num_samples)],
}
production_df = pd.DataFrame(production_data)

# Project 1: Quality
quality_data = {
    'BatchID': product_ids,
    'InspectionDate': [random_date_str() for _ in range(num_samples)],
    'Result': [random.choice(['Pass', 'Fail', 'Conditional']) for _ in range(num_samples)],
    'DefectType': [random.choice(defect_types) for _ in range(num_samples)],
    'Severity': [random.choice(severity_levels) for _ in range(num_samples)],
    'Inspector': [fake.name() for _ in range(num_samples)],
}
quality_df = pd.DataFrame(quality_data)

# Save Project 1 to Excel
with ExcelWriter('project1_data.xlsx', engine='openpyxl') as writer:
    production_df.to_excel(writer, sheet_name='Production', index=False)
    products_df.to_excel(writer, sheet_name='Products', index=False)
    quality_df.to_excel(writer, sheet_name='Quality', index=False)

# Project 2: Sales
sales_data = {
    'InvoiceNo': invoice_ids,
    'SaleDate': [random_date_str() for _ in range(num_samples)],
    'CustomerID': customer_ids,
    'ProductID': product_ids,
    'Quantity': [random.randint(1, 50) for _ in range(num_samples)],
    'UnitPrice(USD)': [round(random.uniform(20, 600), 2) for _ in range(num_samples)],
    'Discount(%)': [round(random.uniform(0, 20), 1) for _ in range(num_samples)],
    'ShippingCost(USD)': [round(random.uniform(5, 50), 2) for _ in range(num_samples)],
    'SalesRep': [fake.name() for _ in range(num_samples)],
}
sales_df = pd.DataFrame(sales_data)

# Project 2: Customers
customers_data = {
    'CustomerID': customer_ids,
    'CustomerName': [fake.company() for _ in range(num_samples)],
    'Category': [random.choice(customer_types) for _ in range(num_samples)],
    'Country': [random.choice(countries) for _ in range(num_samples)],
    'RelationshipStartDate': [random_date(2020, 2024) for _ in range(num_samples)],
}
customers_df = pd.DataFrame(customers_data)

# Project 2: Targets
targets_data = {
    'Month': [random_month_str() for _ in range(num_samples)],
    'ProductCategory': [random.choice(categories) for _ in range(num_samples)],
    'SalesTarget(USD)': [random.randint(10000, 100000) for _ in range(num_samples)],
    'ProductionTarget(units)': [random.randint(500, 5000) for _ in range(num_samples)],
}
targets_df = pd.DataFrame(targets_data)

# Save Project 2 to Excel
with ExcelWriter('project2_data.xlsx', engine='openpyxl') as writer:
    sales_df.to_excel(writer, sheet_name='Sales', index=False)
    customers_df.to_excel(writer, sheet_name='Customers', index=False)
    targets_df.to_excel(writer, sheet_name='Targets', index=False)

# Project 3: Machine Operations
machine_operations_data = {
    'Date': [random_date_str() for _ in range(num_samples)],
    'Hour': [f'{random.randint(0, 23):02d}:00' for _ in range(num_samples)],
    'MachineID': machine_ids,
    'ProductionLine': [random.choice(production_lines) for _ in range(num_samples)],
    'Temperature(°C)': [random.randint(20, 70) for _ in range(num_samples)],
    'Vibration(g)': [round(random.uniform(0.05, 0.25), 2) for _ in range(num_samples)],
    'PowerConsumption(kWh)': [round(random.uniform(100, 500), 2) for _ in range(num_samples)],
    'OperationStatus': [random.choice(statuses) for _ in range(num_samples)],
}
machine_operations_df = pd.DataFrame(machine_operations_data)

# Project 3: Materials Inventory
materials_inventory_data = {
    'Date': [random_date_str() for _ in range(num_samples)],
    'MaterialID': material_ids,
    'MaterialName': [fake.word().capitalize() + ' Material' for _ in range(num_samples)],
    'QuantityOnHand': [round(random.uniform(10, 1000), 2) for _ in range(num_samples)],
    'UnitOfMeasure': [random.choice(units_of_measure) for _ in range(num_samples)],
    'UnitCost': [round(random.uniform(1, 100), 2) for _ in range(num_samples)],
    'MinimumStockLevel': [round(random.uniform(5, 200), 2) for _ in range(num_samples)],
    'LastReceivedDate': [random_date_str() for _ in range(num_samples)],
    'ExpiryDate': [random_date(2025, 2027) if random.random() > 0.3 else '' for _ in range(num_samples)],
    'SupplierID': supplier_ids,
}
materials_inventory_df = pd.DataFrame(materials_inventory_data)

# Project 3: Labor
labor_data = {
    'Date': [random_date_str() for _ in range(num_samples)],
    'DepartmentID': department_ids,
    'ShiftID': shift_ids,
    'EmployeeCount': [random.randint(5, 50) for _ in range(num_samples)],
    'RegularHours': [round(random.uniform(4, 8), 1) for _ in range(num_samples)],
    'OvertimeHours': [round(random.uniform(0, 4), 1) for _ in range(num_samples)],
    'AbsenteeHours': [round(random.uniform(0, 2), 1) for _ in range(num_samples)],
    'LaborCost': [round(random.uniform(500, 5000), 2) for _ in range(num_samples)],
    'ProductivityIndex': [round(random.uniform(0.5, 1.5), 2) for _ in range(num_samples)],
}
labor_df = pd.DataFrame(labor_data)

# Project 3: Maintenance
maintenance_data = {
    'MaintenanceID': [f'MNT-{1000 + i}' for i in range(num_samples)],
    'MachineID': machine_ids,
    'MaintenanceType': [random.choice(maintenance_types) for _ in range(num_samples)],
    'StartDateTime': [random_datetime_str() for _ in range(num_samples)],
    'EndDateTime': [random_datetime_str() for _ in range(num_samples)],
    'TechnicianID': [f'TECH{i % 20 + 1}' for i in range(num_samples)],
    'MaintenanceCost': [round(random.uniform(100, 2000), 2) for _ in range(num_samples)],
    'DowntimeHours': [round(random.uniform(1, 24), 1) for _ in range(num_samples)],
    'PartsReplaced': [fake.word().capitalize() + ' Component' for _ in range(num_samples)],
    'MaintenanceNotes': [fake.sentence() for _ in range(num_samples)],
}
maintenance_df = pd.DataFrame(maintenance_data)

# Project 3: Financial
financial_data = {
    'Date': [random_date_str() for _ in range(num_samples)],
    'TransactionType': [random.choice(transaction_types) for _ in range(num_samples)],
    'DepartmentID': department_ids,
    'ProductLineID': [random.choice(production_lines) for _ in range(num_samples)],
    'Amount': [round(random.uniform(1000, 50000), 2) for _ in range(num_samples)],
    'AccountID': [f'ACC{i % 10 + 1}' for i in range(num_samples)],
    'BudgetCategory': [fake.word().capitalize() for _ in range(num_samples)],
    'ProjectID': [f'PRJ{i % 5 + 1}' for i in range(num_samples)],
}
financial_df = pd.DataFrame(financial_data)

# Project 3: Shipping
shipping_data = {
    'ShipmentID': [f'SHP-{1000 + i}' for i in range(num_samples)],
    'InvoiceID': invoice_ids,
    'ShipDate': [random_date_str  for _ in range(num_samples)],
    'DeliveryDate': [random_date_str() for _ in range(num_samples)],
    'PromisedDate': [random_date_str() for _ in range(num_samples)],
    'ShipmentStatus': [random.choice(shipment_statuses) for _ in range(num_samples)],
    'CarrierID': [f'CAR{i % 10 + 1}' for i in range(num_samples)],
    'ShippingCost': [round(random.uniform(50, 500), 2) for _ in range(num_samples)],
    'Weight': [round(random.uniform(10, 1000), 2) for _ in range(num_samples)],
    'Destination': [fake.city() for _ in range(num_samples)],
}
shipping_df = pd.DataFrame(shipping_data)

# Project 3: Suppliers
suppliers_data = {
    'SupplierID': supplier_ids,
    'SupplierName': [fake.company() for _ in range(num_samples)],
    'MaterialCategory': [fake.word().capitalize() for _ in range(num_samples)],
    'LeadTimeDays': [random.randint(1, 30) for _ in range(num_samples)],
    'QualityRating': [round(random.uniform(1, 10), 1) for _ in range(num_samples)],
    'ContractStart': [random_date(2020, 2024) for _ in range(num_samples)],
    'ContractEnd': [random_date(2025, 2030) for _ in range(num_samples)],
    'City': [fake.city() for _ in range(num_samples)],
    'Country': [fake.country() for _ in range(num_samples)],
    'OnTimeDeliveryRate': [round(random.uniform(50, 100), 1) for _ in range(num_samples)],
}
suppliers_df = pd.DataFrame(suppliers_data)

# Project 3: Energy Consumption
energy_consumption_data = {
    'Date': [random_date_str() for _ in range(num_samples)],
    'FacilityArea': [fake.word().capitalize() + ' Area' for _ in range(num_samples)],
    'EnergyType': [random.choice(energy_types) for _ in range(num_samples)],
    'Consumption': [round(random.uniform(100, 10000), 2) for _ in range(num_samples)],
    'UnitOfMeasure': [random.choice(energy_units) for _ in range(num_samples)],
    'Cost': [round(random.uniform(50, 5000), 2) for _ in range(num_samples)],
    'PeakUsageTime': [random_time_str() for _ in range(num_samples)],
    'TemperatureExternal': [round(random.uniform(-10, 40), 1) for _ in range(num_samples)],
}
energy_consumption_df = pd.DataFrame(energy_consumption_data)

# Save Project 3 to Excel
with ExcelWriter('project3_data.xlsx', engine='openpyxl') as writer:
    machine_operations_df.to_excel(writer, sheet_name='Machine_Operations', index=False)
    materials_inventory_df.to_excel(writer, sheet_name='Materials_Inventory', index=False)
    labor_df.to_excel(writer, sheet_name='Labor', index=False)
    maintenance_df.to_excel(writer, sheet_name='Maintenance', index=False)
    financial_df.to_excel(writer, sheet_name='Financial', index=False)
    shipping_df.to_excel(writer, sheet_name='Shipping', index=False)
    suppliers_df.to_excel(writer, sheet_name='Suppliers', index=False)
    energy_consumption_df.to_excel(writer, sheet_name='Energy_Consumption', index=False)

print("Excel files created: project1_data.xlsx, project2_data.xlsx, project3_data.xlsx")