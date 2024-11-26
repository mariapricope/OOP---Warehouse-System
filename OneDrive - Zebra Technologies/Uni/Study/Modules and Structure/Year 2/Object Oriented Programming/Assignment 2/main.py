#main.py
from warehouse_system import DatabaseManager, Warehouse, ConcreteStockItem

if __name__ == "__main__":
    db_manager = DatabaseManager()
    warehouse = Warehouse(db_manager)

    print("Welcome to the Inventory Management System")

    while True:
        print("\nMenu:")
        print("1. Add Item")
        print("2. Remove Item")
        print("3. View Inventory")
        print("4. Update Item")
        print("5. Exit")
        choice = input("Enter your choice: ")

        if choice == '1':
            name = input("Enter item name: ")
            try:
                quantity_str = input("Enter item quantity: ")
                quantity = int(quantity_str)
                if quantity < 0:
                    raise ValueError("Quantity cannot be negative.")
            except ValueError:
                print("Invalid quantity. Please enter a non-negative integer.")
                continue
            
            category = input("Enter item category: ")
            item = ConcreteStockItem(None, name, quantity, category)
            warehouse.add_item(item)
            print("Item added successfully!")

        elif choice == '2':
            try:
                item_id = int(input("Enter item ID to remove: "))
                
                # Check if the item exists
                item = warehouse.get_item(item_id)
                if item is None:
                    print("Item with the specified ID does not exist.")
                    continue
                
                # Proceed with removal if the item exists
                warehouse.remove_item(item_id)
                print("Item removed successfully!")

            except ValueError:
                print("Invalid ID. Please enter a valid integer.")

        elif choice == '3':
            print("\nCurrent Inventory:")
            items = warehouse.list_items()
            if not items:
                print("No items in inventory.")
            else:
                for item in items:
                    print(item)

        elif choice == '4':
            try:
                item_id = int(input("Enter item ID to update: "))
                name = input("Enter new item name (leave blank to keep current): ")
                quantity_str = input("Enter new item quantity (leave blank to keep current): ")
                category = input("Enter new item category (leave blank to keep current): ")

                # Fetch current item details
                current_item = warehouse.get_item(item_id)
                if not current_item:
                    print("Item not found.")
                    continue

                # Update details
                new_name = name if name else current_item.name
                new_quantity = int(quantity_str) if quantity_str else current_item.quantity
                if quantity_str and new_quantity < 0:
                    print("Quantity cannot be negative.")
                    continue
                new_category = category if category else current_item.category

                # Create updated item
                updated_item = ConcreteStockItem(item_id, new_name, new_quantity, new_category)
                warehouse.update_item(updated_item)
                print("Item updated successfully!")

            except ValueError:
                print("Invalid input. Please enter a valid item ID and quantity.")

        elif choice == '5':
            print("Exiting the system. Goodbye!")
            break

        else:
            print("Invalid choice. Please try again.")

    db_manager.close()
