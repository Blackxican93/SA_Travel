#!/usr/bin/python3
# Required imports
import pandas as pd

def main():
    # Define xlsx file
    excel_file = 'SA_Travel.xlsx'

    # Define sheets and ensure they are indexed
    # df = pd.read_excel('SA_Travel.xlsx', sheet_name = [0,1,2,3,4,5])
    travel_sheet0 = pd.read_excel(excel_file, sheet_name=0, index_col=0)
    travel_sheet1 = pd.read_excel(excel_file, sheet_name=1, index_col=0)
    travel_sheet2 = pd.read_excel(excel_file, sheet_name=2, index_col=0)
    travel_sheet3 = pd.read_excel(excel_file, sheet_name=3, index_col=0)
    travel_sheet4 = pd.read_excel(excel_file, sheet_name=4, index_col=0)
    travel_sheet5 = pd.read_excel(excel_file, sheet_name=5, index_col=0)

    another_topic = "Y"
    # While loop required to ask user if they want to explore more
    while another_topic == "Y":
        # Promt user to select a topic they would like to explore
        print("So I hear you are visiting San Antonio, TX? Awesome!")
        print("Based on some research from TripAdvisor, I'll give you some top destinations based on topics you may want to explore.")

        topics = ["0-Tours & Sightseeing", "1-Outdoors", "2-Shopping", "3-Historic", "4-Amusement & Museums", "5-Great Views"]
        print(topics)

        user_selection = int(input("Which topic would you like to explore? Pick a number.  "))
        print("user_selection:", user_selection)

        if user_selection == 0:
            print(travel_sheet0.head())
        elif user_selection == 1:
            print(travel_sheet1.head())
        elif user_selection == 2:
            print(travel_sheet2.head())
        elif user_selection == 3:
            print(travel_sheet3.head())
        elif user_selection == 4:
            print(travel_sheet4.head())
        elif user_selection == 5:
            print(travel_sheet5.head())
        else:
            print("Invalid response")

        # Prompt the user if they want to choose another topic
        another_topic = (input("Would you like to explore another topic (Y/N)? "))    

        if another_topic == "Y":
            print("Sounds good")
        elif another_topic == "N":
            print("Goodbye, have a nice day. ")
        else:
            print("Quitting SA_Travel")

        # DF has 5 rows and 24 columns (indexed by number
        # Export topics from the top dataframe to excel
        travel_sheet0.head().to_excel("Tours_and_Sightseeing.xlsx")
        travel_sheet1.head().to_excel("Outdoors.xlsx")
        travel_sheet2.head().to_excel("Shopping.xlsx")
        travel_sheet3.head().to_excel("Historic.xlsx")
        travel_sheet4.head().to_excel("Amusement_and_Museums.xlsx")
        travel_sheet5.head().to_excel("Great_Views.xlsx")

        # Export topics from the top of the dataframe to json
        travel_sheet0.head().to_json("Tours_and_Sightseeing.json")
        travel_sheet1.head().to_json("Outdoors.json")
        travel_sheet2.head().to_json("Shopping.json")
        travel_sheet3.head().to_json("Historic.json")
        travel_sheet4.head().to_json("Amusement_and_Museums.json")
        travel_sheet5.head().to_json("Great_Views.json")

        # Export topics from the top of the dataframe to csv
        travel_sheet0.head().to_csv("Tours_and_Sightseeing.csv")
        travel_sheet1.head().to_csv("Outdoors.csv")
        travel_sheet2.head().to_csv("Shopping.csv")
        travel_sheet3.head().to_csv("Historic.csv")
        travel_sheet4.head().to_csv("Amusement_and_Museums.csv")
        travel_sheet5.head().to_csv("Great_Views.csv")


if __name__ == "__main__":
    main()
