import re
import Levenshtein
from nltk.tokenize import word_tokenize
from itertools import permutations, combinations

# Sample email text
email_text = """
2 cases Linguine
1 Fettuccine
3 sheets Lasagne (extra thin!!!!)

Thanks. Please deliver to 166 culverview lane, Branchville, NJ, 07826
"""

# Known product list (converted to lowercase)
product_list = ["linguine", "fettuccine", "lasagne", "spaghetti", "penne", "extra thin lasagne sheets"]

# Function to find the closest product match using Levenshtein distance
def closest_product(phrase, product_list, threshold=2):
    closest_match = None
    min_distance = float('inf')
    max_length = 0
    
    for product in product_list:
        distance = Levenshtein.distance(phrase, product)
        #print(f"Comparing '{phrase}' to '{product}' - Distance: {distance}")  # Debug statement
        if distance <= threshold:
            # Prioritize longer, more complex product names
            if (distance < min_distance) or (distance == min_distance and len(product) > max_length):
                min_distance = distance
                max_length = len(product)
                closest_match = product
                #print(f"New closest match: {closest_match} with distance {min_distance} and length {max_length}")  # Debug statement
                
    return closest_match

# Function to generate all permutations of words in a list
def generate_permutations(words):
    all_permutations = []
    for i in range(len(words), 0, -1):  # Start with the longest set and go to smaller sets
        for comb in combinations(words, i):
            all_permutations.extend(permutations(comb))
    return all_permutations

# Function to clean a line by removing non-letter characters except spaces and numbers
def clean_line(line):
    return re.sub(r'[^a-zA-Z0-9 ]', '', line).lower()

# Function to process each line and extract order details
def process_line(line):
    cleaned_line = clean_line(line)
    #print(f"Processing line: {cleaned_line}")  # Debug statement
    words = word_tokenize(cleaned_line)
    count = None
    product = None
    
    for i, word in enumerate(words):
        if word.isdigit():
            count = int(word)
            #print(f"Found count: {count}")  # Debug statement
            potential_product_words = words[i+1:]
            #print(f"Potential product words: {potential_product_words}")  # Debug statement
            all_permutations = generate_permutations(potential_product_words)
            for permutation in all_permutations:
                permutation_phrase = " ".join(permutation)
                #print(f"Testing permutation: {permutation_phrase}")  # Debug statement
                product = closest_product(permutation_phrase, product_list)
                if product:
                    #print(f"Match found: {product}")  # Debug statement
                    return count, product  # Exit as soon as a match is found
            break
    
    if count is None:
        # Assume 1 case if no count is provided
        count = 1
        #print(f"Default count: {count}")  # Debug statement
        all_permutations = generate_permutations(words)
        for permutation in all_permutations:
            permutation_phrase = " ".join(permutation)
            #print(f"Testing permutation: {permutation_phrase}")  # Debug statement
            product = closest_product(permutation_phrase, product_list)
            if product:
                #print(f"Match found: {product}")  # Debug statement
                return count, product  # Exit as soon as a match is found
    
    return count, product

# Main function to extract orders from email text
def extract_orders(email_text):
    orders = []
    lines = email_text.strip().split('\n')
    
    for line in lines:
        if line.strip():  # Skip empty lines
            count, product = process_line(line)
            if product:
                orders.append((count, product))
            #input("Press Enter to continue...")  # Pause for debugging
    
    return orders

# Extract orders from the email text
orders = extract_orders(email_text)
print("Orders:", orders)
