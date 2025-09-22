# ProdGen

ProdGen is a desktop application for managing product categories, brands, and models, and for generating listings based on customizable templates.

## Manual QA

Follow the steps below to confirm that duplicate rename attempts are handled gracefully:

1. Launch the application with `python main.py` and ensure at least two categories exist (for example, "Смартфони" and "Планшети").
2. Attempt to rename one category to match the other by selecting it, entering the duplicate name, and clicking **Перейменувати** or by double-clicking the tree row and confirming the inline rename.
3. Observe that an error dialog appears with the message about the name already existing and the UI remains responsive.
4. Repeat the same process for a brand and a model inside a category to confirm the same warning dialog appears and the tables continue to function normally after dismissing the message.

These steps verify that the GUI displays a friendly warning instead of crashing when the database rejects a rename because the name already exists.
