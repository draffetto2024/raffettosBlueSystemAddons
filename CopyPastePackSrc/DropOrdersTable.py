import sqlite3
import os


# Get the directory where the script is located
script_dir = os.path.dirname(os.path.abspath(__file__))
db_path = os.path.join(script_dir, 'CopyPastePack.db')

def drop_orders_table():
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # Drop the orders table
        cursor.execute('DROP TABLE IF EXISTS orders')

        cursor.execute('DROP TABLE IF EXISTS ordersandtimestampsonly')
        
        conn.commit()
        print("Orders table successfully dropped")
        
    except sqlite3.Error as e:
        print(f"Error dropping table: {e}")
        
    finally:
        conn.close()

if __name__ == "__main__":
    confirm = input("Are you sure you want to drop the orders table? (y/n): ")
    if confirm.lower() == 'y':
        drop_orders_table()
    else:
        print("Operation cancelled")