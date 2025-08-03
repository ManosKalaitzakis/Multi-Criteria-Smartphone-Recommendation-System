# Multi-Criteria-Smartphone-Recommendation-System
A complete platform for recommending smartphones tailored to the Greek market. It combines web scraping for data collection, topic modeling (LDA) to analyze reviews, and multi-criteria decision-making (UTA methods) to rank and suggest the best smartphones based on multiple factors and user preferences.

This project implements a **smartphone recommendation platform** using data scraping, topic modeling (LDA), and multi-criteria decision-making (MCDA), tailored for the Greek market.

---

## 🔧 Components

### 1. DataRetriever - Web Scraping Engine
Extracts specifications and reviews for smartphones from Greek e-commerce sources.

- **Language**: Python + Selenium + BeautifulSoup  
- **Input**: UrlKinhtwn.txt (list of product URLs)  
- **Output**: kinhta25.xlsx with specs and review data  
- **Features**:  
  - Custom parser for RAM, CPU, display, etc.  
  - Dual-sheet Excel output  

---

### 2. LDA - Topic Modeling of Reviews
Analyzes review texts to extract common themes using LDA (Latent Dirichlet Allocation).

- **Libraries**: Gensim, NLTK, pyLDAvis  
- **Data**: kinhta25.xlsx, stopwords_el.txt  
- **Output**:  
  - LDA_clusters.xlsx: Topic assignments  
  - Visualizations via pyLDAvis  
- **Custom logic**:  
  - Uses MALLET backend for improved LDA  
  - Levenshtein distance for similarity detection  

---

### 3. Agent_Allocator - Personalized Recommendation Engine
Allocates smartphones based on user preferences using weighted criteria.

- **Excel Files**:
  - Apaithseis_Xrhsth.xlsx: User preferences  
  - kritiria.xlsx: Decision criteria weights  
  - xarakthristika_kinhtwn.xlsx: All phone specs  
  - dedomena_apokleismou.xlsx: Filters  
- **Logic**:
  - Reads from Excel sheets
  - Computes scores based on selected features
  - Applies brand filters
- **Output**: Ranked list of phones best matching user profile

---

### 4. Utastar Utilities Calculator - Multi-Criteria Evaluation with UTA Methods
Implements UTA, UTASTAR, and UTADIS decision-making algorithms.

- **Main script**: UtastarUtilitiesCalcuator.py  
- **Helper modules**:
  - utamethods.py: Base utility functions  
  - utautastarfunctions.py: UTASTAR logic  
  - utadisfunctions.py: UTADIS logic  
  - filters.py: Constraint filters  
- **Notebook**: MultiCrit-tutorial.ipynb (example/tutorial)  
- **Inputs**:
  - xarakthristika_kinhtwn.xlsx: Specs  
  - bathmologies_kinhtwn_sunolou_agoras.xlsx: Full market scores  
  - bathmologies_kinhtwn_sunolou_anaforas.xlsx: Reference scores  
- **Output**: Results.xlsx (utility values and rankings)

---

## 📊 Data Flow Diagram

![Data Flow](Διάγραμμα%20Δεδομένων.png)

---

## 🗂️ Folder Structure

.
├── DataRetriever/
├── LDA/
├── Agent_Allocator/
├── UtastarUtilitiesCalcuator/
├── Διάγραμμα Δεδομένων.png
