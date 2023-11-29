import os
from dotenv import load_dotenv
import pyNote
import matplotlib.pyplot as plt

# Load environment variables from .env file
load_dotenv(".env")

# Read credentials from environment variables
client_id = os.getenv("CLIENT_ID")
tenant_id = os.getenv("TENANT_ID")

# Create a new notebook or access an existing one
nb = pyNote.NoteBook(
    name="Sadman Ahmed Shanto @ Work",
    page="Test",
    title="test 2",
)

nb.print("test 1")

# Create a sample plot
x = [0, 1, 2, 3, 4]
y = [0, 1, 4, 9, 16]
plt.plot(x, y)
plt.xlabel("x")
plt.ylabel("y")

# Save the plot to OneNote
nb.savefig(plt.gcf(), "x_y.png")

plt.show()

