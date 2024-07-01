import os
import glob

# Define the patterns to match
patterns = ['*(1)*', '*(2)*', '*(3)*', '*(4)*', '*(5)*']

# Loop through each pattern
for pattern in patterns:
    # Use glob to find files matching the pattern
    for filename in glob.glob(pattern):
        # Use os.remove to delete the file
        os.remove(filename)