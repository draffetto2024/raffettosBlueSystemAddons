import sqlite3
import pandas as pd

def show_database(db_path):
    # Connect to the database
    conn = sqlite3.connect(db_path)
    
    # Query the database
    query = "SELECT * FROM orders"
    df = pd.read_sql_query(query, conn)
    
    # Close the connection
    conn.close()
    
    # Display the data
    print(df)

# Show the database contents
show_database('orders.db')