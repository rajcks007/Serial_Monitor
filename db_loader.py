import pandas as pd
import mysql.connector
import os
import sys

TABLE_NAME = 'product_station3'
CSV_PATH = 'AM60.csv' # Set path here

def load_csv_to_db():
    # Check if CSV file exists
    if not os.path.exists(CSV_PATH):
        print(f"CSV file not found: {CSV_PATH}")
        exit()
    full_df = pd.read_csv(CSV_PATH)
    df = full_df.iloc[[0]]  # Only keep the second row (index 0 after skipping header)
    print(df)

    # Connect to MySQL
    try:
        conn = mysql.connector.connect(
            host = "gitserver.local",
            port = 3306,
            user = "lol",
            passwd = "not try to find it",
            database = "Fast_Product"
            )
        cursor = conn.cursor()
    except mysql.connector.Error as err:
        return False, f"Database connection failed: {err}"

    try:
        # Prepare column names and placeholders
        columns = ', '.join(df.columns)
        placeholders = ', '.join(['%s'] * len(df.columns))

        """
    The columns and placeholders are dynamically generated based on the CSV file.
    # Manually define the placeholders and columns to match the CSV file
    columns = (
        "Timestamp, Serial Number, LED AVERAGE, VCC AVERAGE, PLUS AVERAGE, MINUS AVERAGE,\
              BATTERY AVERAGE, LED, VCC, PLUS VOLT, MINUS VOLT, BATTERY VOLT, RASPBERRY PI, \
                RASPBERRY PI RUN, SCREEN_1, SCREEN_2, SCREEN_3, SCREEN_4, SCREEN_5, SCREEN_6"
    )
    
    placeholders = "%s, %s, %s, %s, %s, %s, %s, %s, %s, \
         %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s"

    # The above is a manual definition of the columns and placeholders.
    """
    
        # Insert the single row
        row = [str(x) for x in df.iloc[0]]
        cursor.execute(
            f'INSERT INTO {TABLE_NAME} ({columns.lower()}) VALUES ({placeholders})',
            row
        )

        conn.commit()
    except mysql.connector.Error as err:
        return False, f"Error inserting data: {err}"
    finally:
        # Close the cursor and connection
        cursor.close()
        conn.close()
        del cursor
        del conn