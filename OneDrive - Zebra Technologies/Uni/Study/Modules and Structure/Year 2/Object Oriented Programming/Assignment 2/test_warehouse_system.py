import unittest
from warehouse_system import DatabaseManager, Warehouse, ConcreteStockItem

class TestWarehouseSystem(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        """Set up a test database for all tests."""
        print("Setting up test database...")
        cls.db_manager = DatabaseManager(db_name="test_warehouse.db")
        cls.warehouse = Warehouse(cls.db_manager)

    @classmethod
    def tearDownClass(cls):
        """Clean up the test database after all tests."""
        print("Tearing down test database...")
        cls.db_manager.cursor.execute("DROP TABLE IF EXISTS stock_items")
        cls.db_manager.connection.close()

    def setUp(self):
        """Clear the warehouse and database before each test."""
        print("Clearing database and resetting warehouse state...")
        self.db_manager.cursor.execute("DELETE FROM stock_items")
        self.db_manager.connection.commit()

    def test_add_item(self):
        print("Running test_add_item...")
        item = ConcreteStockItem(None, "Box", 15, "Packaging")
        self.warehouse.add_item(item)
        items = self.warehouse.list_items()
        print(f"Items in warehouse after add: {items}")
        self.assertEqual(len(items), 1)
        self.assertEqual(items[0].name, "Box")

    def test_remove_item(self):
        print("Running test_remove_item...")
        item = ConcreteStockItem(None, "Core", 20, "Materials")
        self.warehouse.add_item(item)
        self.warehouse.remove_item(item.item_id)
        items = self.warehouse.list_items()
        print(f"Items in warehouse after remove: {items}")
        self.assertEqual(len(items), 0)

    def test_list_items(self):
        print("Running test_list_items...")
        item1 = ConcreteStockItem(None, "Material", 30, "Raw")
        item2 = ConcreteStockItem(None, "Toolbox", 5, "Tools")
        self.warehouse.add_item(item1)
        self.warehouse.add_item(item2)
        items = self.warehouse.list_items()
        print(f"Listed items: {items}")
        self.assertEqual(len(items), 2)

    def test_update_quantity(self):
        print("Running test_update_quantity...")
        item = ConcreteStockItem(None, "Pallet", 10, "Logistics")
        self.warehouse.add_item(item)
        item.update_quantity(15)
        print(f"Item quantity after update: {item.quantity}")
        self.assertEqual(item.quantity, 25)
        print("Testing invalid quantity update...")
        with self.assertRaises(ValueError):
            item.update_quantity(-50)

    def test_database_integration(self):
        print("Running test_database_integration...")
        item = ConcreteStockItem(None, "Crate", 40, "Storage")
        self.warehouse.add_item(item)
        items = self.warehouse.list_items()
        print(f"Items in database after add: {items}")

        # Simulate a new session with the database
        new_warehouse = Warehouse(self.db_manager)
        items_in_new_session = new_warehouse.list_items()
        print(f"Items in new warehouse instance: {items_in_new_session}")
        self.assertEqual(len(items_in_new_session), 1)
        self.assertEqual(items_in_new_session[0].name, "Crate")

    def test_missing_values(self):
        print("Running test_missing_values...")
        # Test missing name
        with self.assertRaises(ValueError):
            ConcreteStockItem(None, "", 10, "Category")  # Missing name

        # Test missing category
        with self.assertRaises(ValueError):
            ConcreteStockItem(None, "ItemName", 10, "")  # Missing category

        # Test negative quantity
        with self.assertRaises(ValueError):
            ConcreteStockItem(None, "ItemName", -10, "Category")  # Negative quantity

if __name__ == "__main__":
    unittest.main()
