# Faraday Grid Wholesale market model

Harry van der Weijde

h.vanderweijde@ed.ac.uk

For installation and usage instructions: see documentation file

Latest edit:
31/1/18 Resolved bug that led to renenewable output data being pulled from the first column of raw data for all renewable generators. Fix involved creating a new set with only renewable generators, and recoding the loop that generates the availability dictionary to first fill it with 1s, and then replace these for the renewable generators.

