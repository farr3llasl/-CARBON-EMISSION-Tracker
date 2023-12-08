import pandas as pd
from tkinter import *
from tkinter import ttk
from tkinter.messagebox import showinfo
import matplotlib.pyplot as plt
from urllib.request import urlretrieve

class MainWindow:
   def __init__(self, root):
       root.title("Main Window")
       root.geometry("800x150")


       # Main Frame
       self.frame = ttk.Frame(root, padding="10 10 10 10")
       self.frame.grid(column=0, row=0, sticky=(N, W, E, S))
       root.columnconfigure(0, weight=1)
       root.rowconfigure(0, weight=1)


       # URL Entry and Get Data Button
       self.url_var = StringVar()
       self.url_entry = ttk.Entry(self.frame, textvariable=self.url_var)
       self.url_entry.grid(column=0, row=0, padx=10, pady=10, sticky=(N, W, E, S), columnspan=3)


       self.get_data_button = ttk.Button(self.frame, text="Get Data", command=self.fetch_data)
       self.get_data_button.grid(column=3, row=0, padx=10, pady=10, sticky=(N, W, E, S))


       # Dropdown Menus - Country1, Country2, Sector
       self.country_1_var = StringVar()
       self.country_2_var = StringVar()
       self.sector_var = StringVar()


       self.country_1_dropdown = ttk.Combobox(self.frame, textvariable=self.country_1_var, values=[])
       self.country_2_dropdown = ttk.Combobox(self.frame, textvariable=self.country_2_var, values=[])
       self.sector_dropdown = ttk.Combobox(self.frame, textvariable=self.sector_var, values=[])


       self.country_1_dropdown.grid(column=0, row=2, padx=10, pady=10)
       self.country_2_dropdown.grid(column=1, row=2, padx=10, pady=10)
       self.sector_dropdown.grid(column=2, row=2, padx=10, pady=10)


       # Graph Button
       self.graph_button = ttk.Button(self.frame, text="Graph", command=self.show_graph)
       self.graph_button.grid(column=3, row=2, padx=10, pady=10, columnspan=2, sticky=(N, W, E, S))


   def fetch_data(self):
       try:
           file_name = 'global_emissions3.xlsx'
           urlretrieve(self.url_entry.get(), file_name)
           #print("Ha")
           # Read the Excel file using pandas
           data = pd.read_excel(file_name, sheet_name='fossil_CO2_by_sector_country_su')
           self.country_1_dropdown['values'] = data['Country'].tolist()
           # print(data['Country'])
           self.country_2_dropdown['values'] = data['Country'].tolist()
           self.sector_dropdown['values'] = list(set(data['Sector'].tolist()))
           # print(data['1970'])
           return data
       except Exception as e:
           showinfo("Get Data", f"Error fetching data: {str(e)}")
           return None


   def show_graph(self):
       try:
           # Fetch data using the fetch_data function
           data = self.fetch_data()
           country_1 = self.country_1_var.get()
           country_2 = self.country_2_var.get()
           sector = self.sector_var.get()
           print(country_1, country_2, sector)
           subset_countries = [country_1, country_2]
           subset_countries_years = []


           for country in subset_countries:
               subset_countries_years.extend(
                   [country + ' (1990)', country + ' (2000)', country + ' (2005)', country + ' (2015)',
                    country + ' (2020)', country + ' (2021)', country + ' (2022)'])
           print(subset_countries_years)


           sector_indices = data.index[data['Sector'] == sector].tolist()
           print(sector_indices)
           s_data = data.loc[sector_indices[0]: sector_indices[-3], :]
           print(s_data)


           index1 = s_data['Country'].tolist().index(country_1)
           index2 = s_data['Country'].tolist().index(country_2)
           # print(index1)
           # print(s_data.iloc[index1][1990.0])


           subset_emissions = [s_data.iloc[index1][1990], s_data.iloc[index1][2000], s_data.iloc[index1][2005],
                               s_data.iloc[index1][2015], s_data.iloc[index1][2020], s_data.iloc[index1][2021],
                               s_data.iloc[index1][2022], s_data.iloc[index2][1990], s_data.iloc[index2][2000],
                               s_data.iloc[index2][2005], s_data.iloc[index2][2015], s_data.iloc[index2][2020],
                               s_data.iloc[index2][2021], s_data.iloc[index2][2022]]
           # print(subset_emissions)


           plt.clf()
           plt.barh(subset_countries_years, subset_emissions)
           plt.title("Carbon Emissions, Compared between Countries (Sector: " + sector + ")")
           plt.xlabel("Emissions (Metric Tons CO2)")
           plt.ylabel("Country")
           plt.tight_layout()
           plt.show()


       except Exception as e:
           showinfo("Graph", f"Error generating graph: {str(e)}")




if __name__ == "__main__":
   root = Tk()
   MainWindow(root)
   root.mainloop()