import csv

# Open the CSV file
with open('Tienda.csv', 'r') as csvfile:

    reader = csv.reader(csvfile, delimiter=';')

    # Iterate over each row in the CSV file
    for row in reader:
        print(row[1])
