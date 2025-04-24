import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# --- Classification Function ---
def classify_with_labels(documents, sentiment_criteria):
    vectorizer = TfidfVectorizer()
    vectorized_texts = vectorizer.fit_transform(sentiment_criteria + documents)

    category_vectors = vectorized_texts[:len(sentiment_criteria)]
    document_vectors = vectorized_texts[len(sentiment_criteria):]
    cosine_similarities = cosine_similarity(document_vectors, category_vectors)

    labels = []
    for i in range(len(documents)):
        best_match_idx = np.argmax(cosine_similarities[i])
        best_category = sentiment_criteria[best_match_idx]
        labels.append(best_category)
    return labels

# --- Sentiment Criteria ---
sentiment_criteria = [
    "Order Backlog", "Revenue Growth", "Net Income", "Operating Income", "Gross Profit",
    "Product Demand", "Awarded Contracts", "Market Expansion", "Revenue Diversification", "Market Share",
    "Research and Development", "Trade Payables", "Contract Liabilities", "Strategic Shifts",
    "Restructuring Losses", "Cash Flow Decrease", "Operating Expenses", "Intersegment Revenue",
    "Trade and Unbilled Receivables", "Supply Chain Disruptions", "Geopolitical Impact",
    "Production Relocation", "Stock-Based Compensation", "Forward-Looking Statements", "Financial Reporting"
]

# --- Load Data ---
df = pd.read_csv("analyst_ratings_processed.csv")

# --- Preprocess ---
df = df.rename(columns={'title': 'sentence'})  # Ensure we use a consistent column name
df = df.dropna(subset=['sentence'])            # Remove rows with missing sentences
news_subset = df['sentence'].tolist()  # Limit for efficiency

# --- Classify Each Sentence ---
print("Classifying news articles...")
classified_labels = classify_with_labels(news_subset, sentiment_criteria)

# --- Add Classification to Original DataFrame ---
df_subset = df.copy()
df_subset['classification'] = classified_labels

# --- Save to CSV ---
output_file = "classified_news_articles.csv"
df_subset.to_csv(output_file, index=False)
print(f"\n✅ Saved classified results to '{output_file}'")

# --- Optional: Show Some Results ---
print(f"\nSentences classified under 'Revenue Growth':")
for sentence in df_subset[df_subset['classification'] == "Revenue Growth"]['sentence'].head(10):
    print(f"  - {sentence}")

# --- Distribution Plot ---
distribution = df_subset['classification'].value_counts()

plt.figure(figsize=(12, 6))
plt.barh(distribution.index, distribution.values, color='skyblue')
plt.xlabel("Number of Articles")
plt.ylabel("Financial Sentiment Category")
plt.title("Distribution of Classified Stock News Articles")
plt.tight_layout()
plt.show()
