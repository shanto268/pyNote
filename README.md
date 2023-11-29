# PyNote

PyNote is a Python library that enables you to interact with OneNote using the OneNote API. You can use this library to create new notebooks, add content, and save Matplotlib figures directly to OneNote.

## Installation

To install the required packages, run:

```bash
pip install python-dotenv msal requests matplotlib
```

## Usage

Below is an example of how to use the PyNote library to create a notebook, add content, and save a Matplotlib figure directly to OneNote.

```python
import matplotlib.pyplot as plt
import numpy as np
from pyNote import NoteBook

# Create a Notebook instance
nb = NoteBook(
    name="test",
    page="page 1",
    title="test 2"
)

# Add text content
nb.print("This is a sample text content.")

# Create a simple plot
x = np.linspace(0, 10, 100)
y = np.sin(x)

plt.plot(x, y)
plt.xlabel("x")
plt.ylabel("y")

# Save the figure to OneNote
nb.savefig("sin_plot.png")

# Display the plot
plt.show()
```

## Setup
