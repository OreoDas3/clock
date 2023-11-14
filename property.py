import random
import uuid

from faker import Faker
import pandas as pd
from datetime import datetime, timedelta

fake = Faker()
insurance_companies='ABC Insurance'

# Function to generate and save a batch of data

class PolicyNumberGenerator:
    def __init__(self):
        self.current_alpha1 = 0
        self.current_alpha2 = 0
        self.current_numeric1 = 0
        self.current_alpha3 = 0
        self.current_numeric2 = 0
        self.current_alpha1_c = 0
        self.current_alpha2_c = 0
        self.current_numeric1_c = 0
        self.current_alpha3_c = 0
        self.current_numeric2_c = 0

    def generate_policy_number(self):
        policy_number = f"POL{chr(ord('A') + self.current_alpha1)}{chr(ord('A') + self.current_alpha2)}{self.current_numeric1:01}{chr(ord('A') + self.current_alpha3)}{self.current_numeric2:02}"
        self.increment_counters()
        return policy_number
    
    def generate_claim_number(self):
        claim_number = f"CLM{chr(ord('A') + self.current_alpha1_c)}{chr(ord('A') + self.current_alpha2_c)}{self.current_numeric1_c:01}{chr(ord('A') + self.current_alpha3_c)}{self.current_numeric2_c:02}"
        self.increment_counters_c()
        return claim_number        

    def increment_counters(self):
        self.current_numeric2 += 1
        if self.current_numeric2 > 99:
            self.current_numeric2 = 0
            self.current_alpha3 += 1
            if self.current_alpha3 == 26:
                self.current_alpha3 = 0
                self.current_numeric1 += 1
                if self.current_numeric1 > 9:
                    self.current_numeric1 = 0
                    self.current_alpha2 += 1
                    if self.current_alpha2 == 26:
                        self.current_alpha2 = 0
                        self.current_alpha1 += 1
                        if self.current_alpha1 == 26:
                            raise ValueError("Exceeded the maximum limit")
    def increment_counters_c(self):
        self.current_numeric2_c += 1
        if self.current_numeric2_c > 99:
            self.current_numeric2_c = 0
            self.current_alpha3_c += 1
            if self.current_alpha3_c == 26:
                self.current_alpha3_c = 0
                self.current_numeric1_c += 1
                if self.current_numeric1_c > 9:
                    self.current_numeric1_c = 0
                    self.current_alpha2_c += 1
                    if self.current_alpha2_c == 26:
                        self.current_alpha2_c = 0
                        self.current_alpha1_c += 1
                        if self.current_alpha1_c == 26:
                            raise ValueError("Exceeded the maximum limit")
                        
def prop_value(argument):
    val = 0
    cons=0
    prem=0
    per=0
    per_clm = 0
    bd=0
    match argument:
        case "Single-family Home":
            val=random.randint(200, 800)*1000
            prem=round(random.uniform(0.01,0.03)*val/100)*100
            per=random.uniform(0.8,0.95)
            per_clm=random.uniform(0.01,0.05)
            cons='Wood frame and Brick'
            ss=fake.random_element(["Yes", "No"])
            sp=fake.random_element(["Yes", "No"])
            bd=random.randint(3,4)
        case "Condons":
            val=random.randint(150, 600)*1000
            prem = round(random.uniform(0.015, 0.04) * val / 100) * 100
            per = random.uniform(0.8, 0.9)
            per_clm = random.uniform(0.01, 0.04)
            cons =fake.random_element(["Concrete and Steel", "Wood and Steel"])
            ss=fake.random_element(["Yes", "No"])
            sp=fake.random_element(["Yes", "No"])
            bd=random.randint(1,2)
        case "Townhouse":
            val =random.randint(180, 700)*1000
            prem = round(random.uniform(0.01, 0.025) * val / 100) * 100
            per = random.uniform(0.8, 0.93)
            per_clm = random.uniform(0.015, 0.05)
            cons = 'Wood and Steel frame'
            ss=fake.random_element(["Yes", "No"])
            sp=fake.random_element(["Yes", "No"])
            bd=random.randint(2,3)
        case "Apartments":
            val =random.randint(100, 500)*1000
            prem = round(random.uniform(0.01, 0.03) * val / 100) * 100
            per = random.uniform(0.8, 0.9)
            per_clm = random.uniform(0.005, 0.03)
            cons = fake.random_element(["Concrete and Steel", "Wood and Steel"])
            sp='Yes'
            ss = 'Yes'
            bd=random.randint(1,2)
        case "Multi-family Home":
            val =random.randint(300, 3000)*1000
            prem = round(random.uniform(0.015, 0.035) * val / 100) * 100
            per = random.uniform(0.8, 0.98)
            per_clm = random.uniform(0.02, 0.06)
            cons = 'Steel and Concrete'
            sp = 'Yes'
            ss = 'Yes'
            bd=random.randint(2,3)
        case "Estate":
            val =random.randint(1000, 10000)*1000
            prem = round(random.uniform(0.02, 0.05) * val / 100) * 100
            per = random.uniform(0.8, 0.99)
            per_clm = random.uniform(0.02, 0.08)
            cons ='Concrete and Steel'
            sp = 'Yes'
            ss = 'Yes'
            bd=random.randint(4,8)
        case "Farmhouse":
            val =random.randint(250, 3000)*1000
            prem = round(random.uniform(0.015, 0.04) * val / 100) * 100
            per = random.uniform(0.8, 0.9)
            per_clm = random.uniform(0.015, 0.05)
            cons = 'Wood frame and brick'
            sp = 'Yes'
            ss = 'Yes'
            bd=random.randint(3,5)
    tca = round(val * per / 100) * 100
    clm=round(val * per_clm / 100) * 100
    return val,prem,cons,tca,clm,ss,sp,bd

def generate_and_save_batch(start_policy_number, batch_size):
    property_policy_data = []
    property_payment_data = []
    property_claim_data = []
    policy = PolicyNumberGenerator()
    policy_numbers = [policy.generate_policy_number() for _ in range(batch_size)]
    claim_number=[policy.generate_claim_number() for _ in range(batch_size)]

    for i in range(1,batch_size):
        policy_number=policy_numbers[i]
        prem_freq = fake.random_element(["Quarterly", "Half Yearly", "Annually"])
        policy_start_date = fake.date_between_dates(date_start=datetime(2017, 9, 1), date_end=datetime(2022, 8, 1))
        policy_end_date = policy_start_date + timedelta(days=365)
        renewal_start_date = policy_end_date - timedelta(days=random.randint(1, 7))
        renewal_end_date = policy_end_date
        claim_date = policy_end_date + timedelta(days=30)
        cov_ty=random.choice(["Personal Property Coverage","Other Structures Coverage",
                               "Liability Coverage","Medical Payments Coverage",
                               "Additional Living Expenses (ALE) Coverage",
                               "Loss of Use Coverage",
                               "Personal Liability Umbrella",
                               "Scheduled Personal Property",
                               "Flood Insurance","Ordinance or Law Coverage"])


        payment_date = fake.date_between_dates(date_start=renewal_start_date, date_end=policy_end_date)
        payment_method = random.choice(["Debit Card","Electronic Funds Transfer (EFT)","Online Banking",
                                        "Cheque","Bank Transfer","Automatic Payments (AutoPay)",
                                        "Money Order","Mobile Payments"])
        payment_status = ["Paid by Customer","Paid by ABC Insurance"]

        property = fake.random_element(
            ["Single-family Home", "Condons", "Townhouse", "Apartments", "Multi-family Home", "Estate", "Farmhouse"])
        property_value = prop_value(property) #Return
        num_bedrooms = property_value[7]
        num_bathrooms=num_bedrooms-1 if num_bedrooms!=1 else 1
        bank=random.choice(["Wells Fargo","JPMorgan Chase","Bank of America",
                            "Quicken Loans (now Rocket Mortgage)","CitiMortgage (Citibank)","U.S. Bank",
                            "PNC Bank","Truist Financial (formerly SunTrust and BB&T)","TD Bank",
                            "HSBC Bank","Flagstar Bank","Guild Mortgage",
                            "Caliber Home Loans","Freedom Mortgage","LoanDepot",
                            "Guaranteed Rate", "Movement Mortgage","New American Funding",
                            "RoundPoint Mortgage Servicing Corporation","Fifth Third Bank"])
        get_claim = random.choice([True, False])

        if get_claim:
            claim_amount = property_value[4]   # Round to the nearest $100
            coc=''
            cause_of_claim = random.choice(["Fire Damage", "Water Damage", "Theft",
                                            "Vandalism", "Falling Objects", "Explosion",
                                            "Smoke Damage", "Power Surge", "Roof Collapse",
                                            "Pipe Burst", "Hail Damage", "Vehicle Collision",
                                            "Structural Collapse", "Gas Leak",
                                            "Tree Falling", "Burglary", "Sewer Backup",
                                            "Accidental Damage", "Building Collapse", "Riot or Civil Commotion",
                                            "Ice Dam", "Sinkhole", "Maintenance Negligence",
                                            "Defective Plumbing", "Electrical Malfunction", "Animal Damage",
                                            "Freezing Pipes", "Faulty Wiring", "Appliance Malfunction",
                                            "Foundation Issues", "Natural Calamity"])
            if cause_of_claim=="Natural Calamity":
                coc=random.choice(["Wildfire",
                                   "Lightning Strike", "Flood", "Hurricane",
                                   "Typhoon", "Tornado", "Mudslide",
                                   "Landslide"])
                cloc=cause_of_claim+' - '+coc
            elif cause_of_claim=="Animal Damage" and property!="Estate"and property!="Multi-family Home"and property!="Apartments":
                coc=random.choice(["Wild animal attack",
                                    "Rodents chewing on wiring or structures",
                                    "Structural damage caused by termites",
                                    "Damage caused by larger animals like horses or cattle"])
                cloc = cause_of_claim + ' - ' + coc
            else:
                cloc=cause_of_claim+coc
            claim_status = "Approved"
            claim_settlement_date = claim_date + timedelta(days=30)
            claim_number_uuid = uuid.uuid4()
            claim_no = claim_number[i]

        else:
            claim_no=None
            claim_date = None
            claim_amount = None
            cloc = None
            claim_status = None
            claim_settlement_date = None


        city=fake.random_element(["New York", "Los Angeles", "Chicago", "Houston", "Phoenix", "Philadelphia", "San Antonio", "San Diego", "Dallas", "San Jose", "Austin", "Jacksonville", "San Francisco", "Columbus", "Indianapolis", "Fort Worth", "Charlotte", "Seattle", "Denver", "Washington, D.C."])
        # Create data for property policy
        phone_number=fake.phone_number()
        alternate_phone_number=fake.phone_number()
        total_no_of_payment = 1 if prem_freq == "Annually" else 2 if prem_freq == "Half Yearly" else 4
        total_payment=total_no_of_payment*property_value[1]
        policy_data = [policy_number,insurance_companies,policy_start_date,policy_end_date,property_value[1],prem_freq,
                       renewal_start_date, renewal_end_date,cov_ty,property_value[3],fake.first_name(),fake.last_name(),
                       fake.random_element(["Male", "Female"]),phone_number, alternate_phone_number,fake.email(),fake.address(),
                       city,fake.zipcode(),property,random.randint(1, 50),property_value[0],fake.street_address(),
                       property_value[2],num_bedrooms, num_bathrooms,property_value[5], property_value[6],bank]

        # Create data for property payment
        payment_no=str(uuid.uuid4())
        payment_data = [payment_no, policy_number, payment_date,"-","-", total_payment,"-", payment_method, "Paid by Policy Holder "+policy_number]
        

        # Add data to lists
        property_policy_data.append(policy_data)
        property_payment_data.append(payment_data)

        if get_claim:
            # Create data for property claim
            payment_data = [payment_no, policy_number, "-", claim_date,claim_settlement_date,"-",claim_amount, "Cheque",payment_status[1]]
            claim_data = [claim_no,policy_data[0],claim_date, claim_amount, cloc,
                          claim_settlement_date, policy_data[10], policy_data[11], policy_data[12],
                          policy_data[13], policy_data[14], policy_data[15],
                           policy_data[16], policy_data[17], policy_data[18],
                           claim_status]
            property_payment_data.append(payment_data)
            property_claim_data.append(claim_data)

    # Create DataFrames for this batch
    property_policy_df = pd.DataFrame(property_policy_data, columns=["Policy_Number","Insurance_Company","Policy_Start_Date","Policy_End_Date","Premium_Amount","Premium_Frequency","Renewal_Start_Date", "Renewal_End_Date",
                                                                     "Coverage_Type", "Total_Coverage_Amount","First_Name","Last_Name","Gender", "Phone_Number", "Alternate_Phone_No","Email","Address","City","Zip Code",
                                                                  "Property_Type", "Property_Age","Property_Value","Property_address","Construction_Type", "Num_Bedrooms", "Num_Bathrooms", "Security_System",
                                                                  "Swimming_Pool", "Mortgage_Provider"])


    property_payment_df = pd.DataFrame(property_payment_data, columns=["Payment_Id", "Policy_Number", "Payment_Date","Claim_Date","Claim_Settlement_Date",
                                                                     "Payment_Amount","Claim_Amount", "Payment_Method", "Payment_Status"])

    property_claim_df = pd.DataFrame(property_claim_data, columns=["Claim_Number","Policy_Number","Claim_Date", "Claim_Amount", "Cause_Of_Claim",
                                                                  "Claim_Settlement_Date","Claimee_first_name","Claimee_last_name","Claimee_gender",
                                                                 "Claimee_Phone_Number", "Alternate_Phone_No","Claimee_email","Claimee_Address","City","Zip Code",
                                                                 "Claim_Status" ])

    # Save the DataFrames to CSV files for this batch
    property_policy_df.to_csv(f"property_policy_batch.csv", index=False)
    print("property_policy_batch.csv is generated")
    property_payment_df.to_csv(f"property_payment_batch.csv", index=False)
    print("property_payment_batch.csv is generated")
    property_claim_df.to_csv(f"property_claim_batch.csv", index=False)
    print("property_claim_batch.csv is generated")
# Generate and save data in batches of 100,000 records each
batch_size = 500

print('Data generation start')

generate_and_save_batch(1, batch_size)
print('Data generation end')
