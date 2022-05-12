from datetime import datetime
import requests
import streamlit as st
import pandas as pd

# Converts miles to km
def miles_to_km(miles):
    return miles * 1.609344

# Isolates number received from API and transforms it in String
def str_to_int(distance):
    distance = distance.replace(",", "")
    distance = distance.replace(" mi", "")
    return float(distance)

#Class representing an Uber ride 
class Uber_Ride: 

    #Initializes Uber_Ride object
    def __init__(self, user,  project_code, amount, pick_up_time, drop_off_time, date):
        self.user = user
        self.project_code = project_code
        self.amount = amount
        self.pick_up_time = pick_up_time
        self.drop_off_time = drop_off_time
        self.date = date
        self.duration = self.get_duration()
        self.emissions = self.get_co2_emissions()

    # Calculates CO2 emissions based on ride duration.
    def get_co2_emissions(self):
        emissions = self.duration.total_seconds() * 1.45865043125317 * 10 ** -7
        return emissions

    # Calculates ride duration by subtracting drop off time by pick up time and returns the duration
    def get_duration(self) :
        datetime1 = datetime.combine(self.date, self.pick_up_time)
        datetime2 = datetime.combine(self.date, self.drop_off_time)
        duration = datetime2 - datetime1
        return duration

#Class representing a company Flight
class Flight: 

    #Initializes Flight object
    def __init__(self, user, date, origin, destination, project_code):
        self.origin = origin
        self.destination = destination
        self.user = user
        self.date = date
        self.project_code = project_code
        # self.distance = self.get_travelling_distance()
        self.distance = 6
        self.emissions = self.get_co2_emissions()

    # Find distance through Matrix Distance API (Google), giving origin and destination; Returns distance in kms
    def get_travelling_distance(self): 

        url = "https://maps.googleapis.com/maps/api/distancematrix/json?origins=" + self.origin + "&destinations=" + self.destination + "&units=imperial&key=AIzaSyD5Jr5tp8WBipYzjG8Vb6_TYyp0Np2fKBs"

        payload = {}
        headers = {}

        response = requests.request("GET", url, headers=headers, data=payload)

        data = response.json()
        distance_in_miles = data["rows"][0]["elements"][0]["distance"]["text"] 

        distance_in_kms = miles_to_km(str_to_int(distance_in_miles))

        return distance_in_kms

    # Calculates and returns flight's CO2 emissions based on flight distance
    def get_co2_emissions(self):
        return self.distance * 2.99401197 * 10 ** -4  # tonnes of CO2

# Class responsible for the calculation of general monthly costs emissions          
class Calculator:

    #Inicializes obejct Calculator
    def __init__(self):
        self.total = 0  # kg CO2
        self.all_flights = []    # List with all flights
        self.all_ubers = []      # List with all uber rides
        self.flights_total = 0   # Total CO2 emissions consumed by flights 
        self.ubers_total = 0     # Total CO2 emissions consumed by uber rides
        self.employees = 0
        print("NEW CALCULATOR")
        self.get_flights()
        self.get_uber_rides()
        self.projects=[]         # List with all projects with emissions estimated
    
    # Takes all flights from excel file, registers them in the system (a list) and calculates total flight emissions
    def get_flights(self):
        
        flights = pd.read_excel("flights_and_uber_rides.xlsx", sheet_name="Flights", usecols="A:E")

        for flight in flights.itertuples():
            user = flight[1]
            date = flight[2]
            origin = flight[3]
            destination = flight[4]
            project_code = flight[5]
            new_flight = Flight(user,date,origin,destination,project_code)
            self.flights_total += new_flight.emissions
            self.all_flights.append(new_flight)

    # Takes all uber rides from excel file, registers them in the system (a list) and calculates total uber rides emissions
    def get_uber_rides(self):
        rides = pd.read_excel("flights_and_uber_rides.xlsx", sheet_name="Uber Rides", usecols="A:F") 
        for ride in rides.itertuples():
            user = ride[1]
            project_code = ride[2]
            amount = ride[3]
            pick_up_time = ride[4]
            drop_off_time = ride[5]
            date = ride[6]
            new_uber = Uber_Ride(user,project_code,amount,pick_up_time,drop_off_time, date)
            self.ubers_total += new_uber.emissions
            self.all_ubers.append(new_uber)

    # Returns total monthly CO2 emissions
    def get_total(self):
        return self.total

    # Calculates and returns monthly CO2 emissions per employee
    def get_emissions_per_employee(self):
        return self.total / self.employees

    # Receives a specific cost CO2 emissions and adds it to the total
    def add_emission(self, emission):
        self.total += emission

    # Receives a number and sets it as the companys total of employees
    def set_employees(self, number_employees):
        self.employees = number_employees

    # Adds project to the system
    def append_project(self, project):
        self.projects.append(project)

    # Returns a list with all the projects codes that are in the system (strings)
    def get_projects_codes(self):
        codes = []
        for project in self.projects:
            codes.append(project.get_code())
        return codes
    
    # Returns all the projects that are registed in the system
    def get_projects (self):
        return self.projects

# Class responsible for the calculation of projects emissions
class Project_Calculator:

    #Inicializes obeject Project_Calculator
    def __init__(self, project_code, employees, general_calculator, duration):
        self.total = 0  
        self.flights = []                            # List with all flights
        self.ubers = []                              # List with all ubers
        self.flights_total = 0                       # Total CO2 emissions consumed by flights 
        self.ubers_total = 0                         # Total CO2 emissions consumed by uber rides 
        self.employees = employees
        self.project_code = project_code
        self.general_calculator = general_calculator # General Calculator where emissions per employee are taken from
        self.duration = duration                     # project duration in months
        self.get_flights_by_code()
        self.get_ubers_by_code()
        general_calculator.append_project(self)      # Regists project in main calculator (Monthly general costs calculator)

    # Takes all flights from excel file that have this project code, registers them in a list and calculates total flights emissions
    def get_flights_by_code(self):
        flights = pd.read_excel("flights_and_uber_rides.xlsx", sheet_name="Flights", usecols="A:E")
        for flight in flights.itertuples():
            if flight[5] == self.project_code:
                user = flight[1]
                date = flight[2]
                origin = flight[3]
                destination = flight[4]
                project_code = flight[5]
                new_flight = Flight(user,date,origin,destination,project_code)
                self.flights.append(new_flight)
                self.flights_total += new_flight.emissions

    # Takes all uber rides from excel file that have this project code, registers them in a list and calculates total uber rides emissions
    def get_ubers_by_code(self):
        rides = pd.read_excel("flights_and_uber_rides.xlsx", sheet_name="Uber Rides", usecols="A:F")
        #rides = pd.read_excel(uploaded_file, sheet_name="Uber Rides", usecols="A:F")
        for ride in rides.itertuples():
            if ride[2] == self.project_code:
                user = ride[1]
                project_code = ride[2]
                amount = ride[3]
                pick_up_time = ride[4]
                drop_off_time = ride[5]
                date = ride[6]
                new_uber = Uber_Ride(user,project_code,amount,pick_up_time,drop_off_time, date)
                self.ubers.append(new_uber)
                self.ubers_total += new_uber.emissions

    # Calculates and returns total project CO2 emissions
    def calculate_total(self):
        self.total= self.ubers_total + self.flights_total + self.employees * float(self.general_calculator.get_emissions_per_employee()* self.duration)
        return self.total

    # Returns project code
    def get_code(self):
        return self.project_code

# Function that makes streamlit and Calculators and outputs run 
def main():
    
    st.image("logo.png")

    if "calculator" not in st.session_state:
            st.session_state.calculator = Calculator()
 
    menu = ["About","General Calculator","Project Calculator", "Project Suggestions"]
    choice = st.sidebar.selectbox("Menu", menu)

    if choice == "General Calculator":

        st.title("General Calculator")
        st.write(" ")
        st.write("In this page please insert your company general monthly costs.") 
        st.write(" ") 
        st.write("Please take into accoun that the number of employees input corresponds to the number of employees associated with these costs")
        st.subheader("Add number of employees:")
        #number_employees = (st.text_input("Number of Employees:"))
        number_employees = st.slider('Number of Employees:', 0, 1000,250)

        #   - - - - - - - - - -  Electricity

        st.write(" ")
        st.subheader("Add Electricity Costs:")
        electricity = (st.text_input("Electricity Costs:"))

        # .245 USD per kWh
        

        # -------------------  Paper Base Product Costs

        st.write(" ")
        st.subheader("Add Paper Base Product Costs:")
        p_b_p = (st.text_input("Paper Base Product Costs:"))

        # --------------------- Telephones and Phone Calls Costs

        st.write(" ")
        st.subheader("Add Telephones and Phone Calls Costs:")
        t_pc = (st.text_input("Telephones and Phone Calls Costs:"))

        # -------------------------Hotels and Restaurants

        st.write(" ")
        st.subheader("Add Hotels and Restaurants Costs:")
        h_r = (st.text_input("Hotels and Restaurants Costs:"))


        # -----------------------Team Building   project scope

        st.write(" ")
        st.subheader("Add Team Building Costs:")
        t_b = (st.text_input("Team Building Costs:"))

    #--------Output-----

        results_button = st.button("See Results")
        if results_button:
            st.session_state.calculator.set_employees(int(number_employees))

            electricity_result = (0.23 * float(electricity)) / (1000.0 * 0.245)
            paper_result = 0.22 / 1000.0 * float(p_b_p)
            telephone_result = 0.20 / 1000.0 * float(t_pc)
            hotel_result = 0.30 / 1000.0 * float(h_r)
            team_building_result = 0.26 / 1000.0 * float(t_b)
            #data = {'Eletricity': [electricity_result], 'Paper Base Product Costs': [paper_result], 'Hotel and restaurant': [hotel_result], "Telephone": [telephone_result], 'Team buildind' : [team_building_result]}
            #df = pd.DataFrame(data)

            st.session_state.calculator.add_emission(electricity_result)
            st.session_state.calculator.add_emission(paper_result)
            st.session_state.calculator.add_emission(telephone_result)
            st.session_state.calculator.add_emission(hotel_result)
            st.session_state.calculator.add_emission(team_building_result)

            
            
            st.title("Total emission (per tonne)")
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Eletricity", round(electricity_result,4), None)
            col2.metric("Paper base product", round(paper_result,4), None)
            col3.metric("Hotel and restaurant", round(hotel_result,4), None)
            col4.metric("Telephone", round(telephone_result,4), None)
            
            col5, col6, col7  =st.columns(3)
            col5.metric("Team building", round(team_building_result,4), None)
            col6.metric("Total Carbon Emission", round(st.session_state.calculator.get_total(),4), None)
            col7.metric("Total Carbon Emission per Employee", round(st.session_state.calculator.get_emissions_per_employee(),4), None)
            st.write("")
            #st.line_chart(df)
            
            
    elif choice == "Project Calculator":
        st.title("Project Calculator")
        st.write("The following information should be related to a specific project. Please insert the following inputs:")
        st.write(" ")
        
        st.subheader("Project Code:")
        st.write("Please take into account that for security measures each project corresponds to: AAA, BBB, CCC or DDD:")
        #project_code = (st.text_input("Project Code:"))
        project_code = st.radio('Project Code:',     #?
                  ['AAA',
                   'BBB',
                   'CCC',
                   'DDD'])

        st.write(" ")
        st.subheader("Number of Employees in the project:")
        #project_employees = st.text_input("Number of Employees in the project:")
        project_employees = st.slider('Number of Employees in the project:', 0, 100, 24)

        st.write(" ")
        st.subheader("Number of Months expected in:")
        month_duration = st.text_input("Number of Months expected:")

        #get project's uber rides and flights
        st.write(" ")

        project_results_button = st.button("Calculate Project Carbon Emissions")
        if project_results_button:
            project = Project_Calculator(project_code, float(project_employees), st.session_state.calculator, float(month_duration))
            st.title("Total emission (per tonne)")
            col11, col12, col13= st.columns(3)
            col11.metric("Total Uber rides", round(project.ubers_total,4), None)
            col12.metric("Total Flights", round(project.flights_total,4), None)
            col13.metric("Total Project's Carbon Emissions", round(project.calculate_total(),4), None)
            

    elif choice == "Project Suggestions":
        
        code = st.selectbox("Choose Project:", st.session_state.calculator.get_projects_codes())
        project_emission = 0
        for project in st.session_state.calculator.get_projects():
            if code == project.get_code():
                project_emission = project.calculate_total()

        print(project_emission)
        st.title("How will Mckinsey offset its emissions?")

        st.write(" ")
        
        #First Option
        st.image("https://th.bing.com/th/id/R.17cfb93317e67f7c99f4c05d3bbbc43a?rik=I7wrRVgGdAJ%2bRg&riu=http%3a%2f%2fjtechsolar.com%2fwp-content%2fuploads%2f2018%2f01%2fjtech-solar-header-0117.jpg&ehk=tmdIaEUxFKgECUKGFpLFPBdOeIPCjc810VlM3LTyctY%3d&risl=&pid=ImgRaw&r=0")
        st.subheader("**First Option:** Global Portfolio")
        st.write("Select the Global Portfolio and your carbon offsetting will be used to support projects around the World")
        st.write("All offset projects in the Global Portfolio are verified to the Verified Carbon Standard (VCS). Our projects are selected predominantly within developing countries, where they assist carbon emission reduction as well as bringing local community benefits. Examples of projects you will be financing:")
        st.write("1.1	Wind base Power Generation by Panama Wind Energy")
        st.write("1.2	Solar Project by ACME")
        st.write("1.3	Peralta I Wind Project")
        st.metric("To support this prject, you will need to spend", round(0.95*project_emission,4), None)
        
        st.write(" ")
        
        #Second Option 
        st.image("https://www.nestle.com/sites/default/files/2020-12/mucilon-reforestation-brazil-feed.jpg")
        st.subheader("**Second Option:** UK Tree Planting: ")
        st.write("Planting is a great way to help sequester carbon emissions. Through photosynthesis trees absorb carbon dioxide to produce oxygen and wood")
        st.metric("To support this prject, you will need to spend", round(24.10*project_emission,4), None)
        st.write(" ")
        
        #Third Option
        st.image("https://images.squarespace-cdn.com/content/v1/5f3bb6551c12282eeeca0e5e/1597754626686-NAWZK6KLOERE2XM7YN50/ke17ZwdGBToddI8pDm48kFmtOlh3rgiGRX_vVwIWqKUUqsxRUqqbr1mOJYKfIPR7LoDQ9mXPOjoJoqy81S2I8N_N4V1vUb5AoIIIbLZhVYwL8IeDg6_3B-BRuF4nNrNcQkVuAT7tdErd0wQFEGFSnBI6sFqUloI5a_OF2UQ3EP0ORVqgf1_MwwwrF6eGKzWhkBXjoW1fEeHsrwHo7P8_5g/image-asset.jpeg?format=1500w")
        st.subheader("**Third:** Reforestation in Kenya: ")
        st.write("With Carbon Footprint you can support local communities in the Great Rift Valley, Kenya")
        st.write(" ")
        st.metric("To support this prject, you will need to spend (in €):", round(19.47*project_emission,4), None)
        st.write(" ")
        
        #Fourth Option
        
        st.image("https://createyourforest.ca/Files/Pages/header/deforestation-header.jpg")
        st.subheader("**Fourth:** Americas Portfolio: ")       
        st.write(" Select Americas Projects and your carbon offsetting will be used to support projects in the Americas region that also deliver additional social and ecological benefits. Examples of projects you will be financing:")
        st.write("1.1	Portel-Para Reducing Deforestation (REDD)")
        st.write("1.2	Mariposas Hydroelectric Project, in Chile")
        st.write("1.3	Peralta I Wind Power Project")
        st.write(" ")
        st.metric("To support this prject, you will need to spend (in €):",round(11.56*project_emission,4), None)
        st.write(" ")
  
        
        #Fifth Option
        st.image("https://www.avignonesi.it/wp-content/uploads/2019/04/cooking-class.jpg")
        st.subheader("**Fifth:** Community Projects: ")
        st.write(" Community Projects and your carbon offsetting will be used to support projects in developing countries that also deliver additional benefits to local communities. Examples of Projects you will be financing:")
        st.write("1.1	Clean and Efficient Cooking and Heating Project (China) ")
        st.write("1.2	The Breathing Space Improved Cooking Stoves Program (India)")
        st.write("1.3	Borehole Rehabilitation Project in Uganda")
        st.write(" ")
        st.metric("To support this prject, you will need to spend (in €):",round(11.56*project_emission,4), None)
        st.write(" ")

    else:
        st.subheader("On the path to net zero.")
        st.write(" ")
        st.write(" ")
        st.video("https://www.youtube.com/watch?v=Rb4ylhVQo7I")
        st.subheader("About us")
        st.write("The value proposition of our program is to enable McKinsey to calculate the CO2 emissions emitted "
                 "by their offices/employees and propose to offset these with environmentally friendly projects ("
                 "i.e., planting trees), thus becoming carbon neutral.")
        st.write("")
        st.write("We designed a CO2 calculator in which McKinsey should enter data on a monthly basis, after closing "
                 "its accounts. At the end of the year, it will then have access to all carbon emissions emitted "
                 "within the scope of the program.")
        st.write("")
        st.write("Following the computation of the tons emitted, our calculator offers several carbon offset choices. "
                 " The user will be able to choose from five different projects, and after doing so, the calculator "
                 "will also tell him how much capital is required to finance the offsetting of those carbon emissions. ")
        st.write("")
        st.write("**Our scope:** ")
        st.write("")
        st.write("•	Electricity (CO2 tonnes/kWh)")
        st.write("• Flights (CO2 tonnes/Km)")
        st.write("• Uber Rides (CO2 tonnes/min)")
        st.write("• Paper based products (CO2 tonnes/€ Spent)")
        st.write("• Telephone, mobile/cell phone call costs (CO2 tonnes/€ Spent)")
        st.write("• Hotels and restaurants (CO2 tonnes/€ Spent)")
        st.write("• Teambuilding: recreational, cultural and porting activities (CO2 tonnes/€ Spent)")
        st.write(" ")
        st.write(" ")
        st.write(" ")
        st.image("team.png")

main()



