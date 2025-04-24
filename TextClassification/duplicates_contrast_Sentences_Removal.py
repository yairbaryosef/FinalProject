import json
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

def filter_sentences(sentences, similarity_threshold=0.85):
    """
    Removes duplicate or semantically similar sentences using TF-IDF and cosine similarity.

    :param sentences: List of sentences.
    :param similarity_threshold: Threshold for considering sentences as duplicates.
    :return: Filtered list of sentences.
    """
    if not sentences:
        return []

    # Compute TF-IDF vectors
    vectorizer = TfidfVectorizer()
    tfidf_matrix = vectorizer.fit_transform(sentences)

    # Compute cosine similarity matrix
    cosine_sim_matrix = cosine_similarity(tfidf_matrix)

    # Track unique sentences
    unique_sentences = []
    seen_indices = set()

    for i in range(len(sentences)):
        if i in seen_indices:
            continue  # Skip already considered sentences

        unique_sentences.append(sentences[i])

        # Mark similar sentences as seen
        for j in range(i + 1, len(sentences)):
            if j not in seen_indices and cosine_sim_matrix[i, j] > similarity_threshold:
                seen_indices.add(j)

    return unique_sentences
