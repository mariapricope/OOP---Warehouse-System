# warehouse_system.py
import sqlite3
from abc import ABC, abstractmethod

# Database Manager
class DatabaseManager:
    def __init__(self, db_name="warehouse.db"):
        self.db_name = db_name
        self.connection = sqlite3.connect(self.db_name)
        self.cursor = self.connection.cursor()
        self.create_table()

    def create_table(self):
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS stock_items (
                               item_id INTEGER PRIMARY KEY AUTOINCREMENT,
                               name TEXT,
                               quantity INTEGER,
                               category TEXT)''')
        self.connection.commit()

    def add_item_to_db(self, item):
        self.cursor.execute('''INSERT INTO stock_items (name, quantity, category) 
                               VALUES (?, ?, ?)''', (item.name, item.quantity, item.category))
        self.connection.commit()
        item.item_id = self.cursor.lastrowid  # Update the item ID after insertion

    def remove_item_from_db(self, item_id):
        self.cursor.execute('DELETE FROM stock_items WHERE item_id = ?', (item_id,))
        self.connection.commit()

    def get_items_from_db(self):
        self.cursor.execute('SELECT * FROM stock_items')
        rows = self.cursor.fetchall()
        return [ConcreteStockItem(row[0], row[1], row[2], row[3]) for row in rows]
    
    def update_item_in_db(self, item):
        self.cursor.execute('''
            UPDATE stock_items
            SET name = ?, quantity = ?, category = ?
            WHERE item_id = ?
        ''', (item.name, item.quantity, item.category, item.item_id))
        self.connection.commit()

    def close(self):
        self.connection.close()

# Abstract StockItem Class (for abstraction)
class StockItem(ABC):
    def __init__(self, item_id, name, quantity, category):
        self.item_id = item_id
        self.name = name
        self.quantity = quantity
        self.category = category

    @abstractmethod
    def update_quantity(self, amount):
        pass

    def __str__(self):
        return f"Item ID: {self.item_id}, Name: {self.name}, Quantity: {self.quantity}, Category: {self.category}"

# Concrete StockItem Class
class ConcreteStockItem(StockItem):
    def update_quantity(self, amount):
        if amount < 0 and self.quantity + amount < 0:
            raise ValueError("Quantity cannot be negative.")
        self.quantity += amount

# Warehouse Class (for encapsulation)

class Warehouse:
    def __init__(self, db_manager):
        self.db_manager = db_manager

    def add_item(self, item):
        """Add an item to the inventory."""
        self.db_manager.add_item_to_db(item)

    def remove_item(self, item_id):
        """Remove an item from the inventory by its ID."""
        self.db_manager.remove_item_from_db(item_id)

    def list_items(self):
        """List all items in the inventory."""
        return self.db_manager.get_items_from_db()

    def get_item(self, item_id):
        """Get a specific item from the inventory by its ID."""
        # This could be optimized to fetch only one item, but for simplicity, we'll just filter from the list
        items = self.db_manager.get_items_from_db()
        for item in items:
            if item.item_id == item_id:
                return item
        return None

    def update_item(self, updated_item):
        """Update an existing item in the inventory."""
        # You would need to implement an update method in the DatabaseManager
        # Here's a placeholder for how it might look:
        self.db_manager.update_item_in_db(updated_item)
