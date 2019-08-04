import os
import csv

csvpath = os.path.join("pybank/resources", "budget_data.csv")

date = []
profit = []
monthly_changes = []

count = 0
total_profit = 0
total_profit_changes = 0
first_profit = 0


with open(csvpath, newline="") as csvfile:
    csvreader = csv.reader(csvfile, delimiter=",")
    csv_header = next(csvreader)

    for row in csvreader:

        count = count + 1

        date.append(row[0])
        
        profit.append(row[1])
        
        total_profit = total_profit + int(row[1])
        
        final_profit = int(row[1])
        monthly_changes_profit = final_profit - first_profit

        monthly_changes.append(monthly_changes_profit)
        total_profit_changes = total_profit_changes + monthly_changes_profit
        first_profit = final_profit
        average_change_profits = (total_profit_changes / count)

        greatest_decrease_profit = min(monthly_changes)
        greatest_increase_profit = max(monthly_changes)

        greatest_increase_date = date[monthly_changes.index(greatest_increase_profit)]
        greatest_decrease_date = date[monthly_changes.index(greatest_decrease_profit)]

    print("Financial Analysis")
    print("----------------------------")
    print("Total Months: " + str(count))
    print("Total: " + "$" + str(total_profit))
    print("Average  Change: " + "$" + str(int(average_change_profits)))
    print("Greatest Increase in Profits:" + str(greatest_increase_date) + " ($" + str(greatest_increase_profit) + ")")
    print("Greatest Decrease in Profits:"+ str(greatest_decrease_date) + " ($" + str(greatest_decrease_profit) + ")")

with open("budget_data.txt", "w") as text: 
    text.write("Financial Analysis" + "\n")
    text.write("----------------------------"+ "\n")
    text.write("Total Months: " + str(count)+ "\n")
    text.write("Total: " + "$" + str(total_profit)+ "\n")
    text.write("Average  Change: " + "$" + str(int(average_change_profits))+ "\n")
    text.write("Greatest Increase in Profits:" + str(greatest_increase_date) + " ($" + str(greatest_increase_profit) + ")"+ "\n")
    text.write("Greatest Decrease in Profits:"+ str(greatest_decrease_date) + " ($" + str(greatest_decrease_profit) + ")"+ "\n")