#!/usr/bin/python3
import pandas as pd


def main():
    #Opening prompt
    df = pd.read_excel('SA_Travel.xlsx', sheet_name = [0,1,2,3,4,5])
    #pd.set_option("display.max_columns", 10, "display.max_rows", None)

    print("So I hear you are visiting San Antonio, TX? Awesome!")
    print("Based on some research from TripAdvisor, I'll give you some top destinations based on topics you may want to explore.")

    topics = ["0-Tours & Sightseeing", "1-Outdoor", "2-Shopping", "3-Historic", "4-Amusement & Museums", "5-Great Views"]
    print(topics)

    user_selection = int(input("Which topic would you like to explore? Pick a number.  "))
    print("user_selection:", user_selection)

    # Define xls file
    excel_file = 'SA_Travel.xlsx'

    # create a DataFrame (DF) object
    travel = pd.read_excel(excel_file)

    # show the first five rows of our DF
    #print(travel.head())

    # Choose the first column "Title" as
    # index (index=0)
    travel_sheet0 = pd.read_excel(excel_file, sheet_name=0, index_col=0)
    travel_sheet1 = pd.read_excel(excel_file, sheet_name=1, index_col=0)
    travel_sheet2 = pd.read_excel(excel_file, sheet_name=2, index_col=0)
    travel_sheet3 = pd.read_excel(excel_file, sheet_name=3, index_col=0)
    travel_sheet4 = pd.read_excel(excel_file, sheet_name=4, index_col=0)
    travel_sheet5 = pd.read_excel(excel_file, sheet_name=5, index_col=0)

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
    # DF has 5 rows and 24 columns (indexed by title)
    # print the top 10 movies in the dataframe
    #print(travel_sheet1.head())
    # export 5 movies from the top dataframe to excel
    travel_sheet0.head(6).to_excel("Top5SA_Travel.xlsx")

    # export 5 movies from the top of the dataframe to json
    travel_sheet0.head(6).to_json("Top5SA_Travel.json")

    # export 5 movies from the top of the dataframe to csv
    travel_sheet0.head(6).to_csv("Top5SA_Travel.csv")


if __name__ == "__main__":
    main()