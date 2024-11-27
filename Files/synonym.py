import os
from nltk.corpus import wordnet
import nltk

# Ensure that the WordNet dataset is downloaded
nltk.download('wordnet')

def get_synonyms(word):
    """Retrieve synonyms for a given word using WordNet."""
    synonyms = set()
    for syn in wordnet.synsets(word):
        for lemma in syn.lemmas():
            synonyms.add(lemma.name().replace('_', ' '))
    return synonyms

def load_compliant_keywords(file_path):
    """
    Load keywords from a file and expand each with its synonyms.
    Each line in the file should contain a single keyword.
    """
    keywords = set()
    try:
        with open(file_path, 'r') as f:
            for line in f:
                word = line.strip()
                synonyms = get_synonyms(word)
                keywords.add(word)  # Add the original word
                keywords.update(synonyms)  # Add all synonyms
        return keywords
    except FileNotFoundError:
        print(f"[ERROR] File not found: {file_path}")
        return set()

# Example usage
file_path = 'compliant_keywords.txt'  # Ensure this file exists in the same directory
compliant_keywords = load_compliant_keywords(file_path)
print("Expanded compliance keywords:", compliant_keywords)
