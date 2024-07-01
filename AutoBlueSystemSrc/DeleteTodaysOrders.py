import sqlite3
from datetime import datetime

def delete_todays_orders(database_path):
    # Get today's date in the 'yyyy-mm-dd' format
    today = datetime.now().strftime('%Y-%m-%d')

    # Connect to the database
    conn = sqlite3.connect(database_path)
    cursor = conn.cursor()

    # Execute the delete query
    cursor.execute("DELETE FROM orders WHERE DATE(date_generated) = ?", (today,))
    
    # Commit the changes
    conn.commit()

    # Close the connection
    conn.close()

    print(f"Orders generated on {today} have been deleted.")

# Specify your database path
database_path = 'orders.db'  # Replace with your actual database file

# Run the deletion function
delete_todays_orders(database_path)
