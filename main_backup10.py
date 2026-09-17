import streamlit as st
import pandas as pd
import requests
import re
import speech_recognition as sr
from streamlit_mic_recorder import mic_recorder
import base64
from urllib.parse import quote_plus
from pathlib import Path
import folium
from streamlit_folium import st_folium

# =====================================================================
# 1. PAGE CONFIGURATION & STYLING
# =====================================================================
st.set_page_config(
    page_title="Culture DNA — Visual & Geo-Heritage Chatbot",
    page_icon="🧬",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Playfair+Display:wght@700&display=swap');
    
    .main-title {
        font-family: 'Playfair Display', serif;
        font-size: 2.3rem;
        font-weight: 700;
        background: linear-gradient(120deg, #B33927 0%, #E65100 50%, #D84315 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.1rem;
    }
    .sub-title {
        color: #555555;
        font-size: 1.05rem;
        margin-bottom: 1.2rem;
    }
    .state-spotlight-card {
        background: #FFFFFF;
        border: 1px solid #FFE0B2;
        border-radius: 12px;
        padding: 20px 24px;
        margin-bottom: 24px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.05);
    }
    .state-spotlight-title {
        font-size: 1.5rem;
        font-weight: 700;
        color: #BF360C;
        margin-bottom: 12px;
    }
    .map-header {
        font-size: 0.95rem;
        font-weight: 600;
        color: #37474F;
        margin-top: 10px;
        margin-bottom: 4px;
    }
    /* Global readability: light main area uses dark text. Sidebar is styled separately below. */
    html, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"] {
        color: #172b3a !important;
    }
    [data-testid="stMain"] .stMarkdown,
    [data-testid="stMain"] .stMarkdown p,
    [data-testid="stMain"] .stMarkdown li,
    [data-testid="stMain"] .stMarkdown span,
    [data-testid="stMain"] [data-testid="stMarkdownContainer"],
    [data-testid="stMain"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stMain"] [data-testid="stMarkdownContainer"] li {
        color: #172b3a !important;
    }

    /* Sidebar readability: never inherit the main-area dark text rules. */
    [data-testid="stSidebar"],
    [data-testid="stSidebar"] [data-testid="stSidebarContent"],
    [data-testid="stSidebar"] .stMarkdown,
    [data-testid="stSidebar"] .stMarkdown p,
    [data-testid="stSidebar"] .stMarkdown li,
    [data-testid="stSidebar"] .stMarkdown span,
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] li,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] h4,
    [data-testid="stSidebar"] small,
    [data-testid="stSidebar"] .stCaption {
        color: #f5f7fa !important;
    }
    [data-testid="stSidebar"] {
        background: #252631 !important;
    }
    [data-testid="stSidebar"] [data-baseweb="select"] > div {
        background: #0d1118 !important;
        border-color: #3b3f4b !important;
        color: #ffffff !important;
    }
    [data-testid="stSidebar"] [data-baseweb="select"] *,
    [data-testid="stSidebar"] [role="combobox"] {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
    }
    [data-testid="stSidebar"] button {
        color: #f5f7fa !important;
    }
    [data-testid="stSidebar"] hr {
        border-color: #444753 !important;
    }
    [data-testid="stChatMessage"] {
        background: #fffaf3 !important;
        border: 1px solid #ead9c5 !important;
        border-radius: 16px !important;
        padding: 14px 16px !important;
        color: #172b3a !important;
    }
    [data-testid="stChatMessage"] [data-testid="stMarkdownContainer"],
    [data-testid="stChatMessage"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stChatMessage"] [data-testid="stMarkdownContainer"] li,
    [data-testid="stChatMessage"] [data-testid="stMarkdownContainer"] span,
    [data-testid="stChatMessage"] strong {
        color: #172b3a !important;
    }
    [data-testid="stChatInput"] {
        background: #fffaf3 !important;
    }
    [data-testid="stChatInput"] textarea {
        color: #172b3a !important;
        -webkit-text-fill-color: #172b3a !important;
        background: #fff !important;
    }
    [data-testid="stChatInput"] textarea::placeholder {
        color: #66727d !important;
        opacity: 1 !important;
    }
    .stChatMessage {
        border-radius: 14px;
        padding: 12px;
    }
    .voice-box {
        background: #fffaf3;
        border: 1px solid #ead9c5;
        border-radius: 14px;
        padding: 8px 10px;
        margin: 6px 0 10px 0;
    }
    .voice-help {
        color: #53616b !important;
        font-size: 0.78rem;
        margin-top: 2px;
    }
    .home-screen-wrap {
        width: 100%;
        margin: 0 auto 1.5rem auto;
        border-radius: 24px;
        overflow: hidden;
        box-shadow: 0 18px 50px rgba(0,0,0,0.18);
        background: #f7ead8;
    }
    .home-screen-wrap img {
        width: 100%;
        display: block;
    }
</style>
""", unsafe_allow_html=True)

# =====================================================================
# 2. DATA LOADING (Preserves Culture_DNA.xlsx)
# =====================================================================
DATA_FILE = Path(__file__).parent / "Culture_DNA.xlsx"

@st.cache_data
def load_data():
    candidate_paths = [
        DATA_FILE,
        Path.cwd() / "Culture_DNA.xlsx",
        Path(__file__).parent.parent / "Culture_DNA.xlsx",
        Path(r"c:\Users\shivi\OneDrive\Desktop\Culture_DNA.xlsx")
    ]
    for p in candidate_paths:
        if p.exists():
            try:
                loaded = pd.read_excel(p)
                loaded = loaded.fillna("")
                # Ensure image_url column is present in the DataFrame structure
                if "image_url" not in loaded.columns:
                    loaded["image_url"] = ""
                return loaded
            except Exception:
                pass
    st.error("Culture_DNA.xlsx data file not found! Please ensure it is in the project folder.")
    empty_df = pd.DataFrame()
    empty_df["image_url"] = ""
    return empty_df

df = load_data()

# =====================================================================
# 3. 🖼️ CUSTOM IMAGE URL REPOSITORY (ADD YOUR OWN IMAGE LINKS HERE)
# =====================================================================
# 👉 YAHAN AAP APNI IMAGE URLS DAAL SAKTE HAIN:
# Agar aap kisi state ke liye apni manpasand image dikhana chahte hain,
# toh neeche us state ke samne quotes "" ke andar direct link paste karein:
# Example: "Punjab": "https://example.com/my_punjab_photo.jpg"
#
# Agar yahan link nahi hoga aur Excel mein bhi link nahi hoga,
# toh chatbot automatically Wikipedia Open Data se authentic HD photo le aayega.
STATE_CUSTOM_IMAGE_URLS = {
    "1.1 Jammu": "",
    "1.2 Kashmir": "",
    "Himachal Pradesh": "",
    "Punjab": "",
    "Haryana": "",
    "Uttarakhand": "",
    "Chandigarh": "",
    "Rajasthan": "",
    "Gujarat": "",
    "Maharashtra": "",
    "Goa": "",
    "Madhya Pradesh": "",
    "Dadra & Nagar Haveli and Daman & Diu": "",
    "Uttar Pradesh": "",
    "Bihar": "",
    "Jharkhand": "",
    "Chhattisgarh": "",
    "Odisha": "",
    "Ladakh": "",
    "West Bengal": "",
    "Sikkim": "",
    "Assam": "",
    "Arunachal Pradesh": "",
    "Meghalaya": "",
    "Andaman & Nicobar Islands": "",
    "Nagaland": "",
    "Manipur": "",
    "Mizoram": "",
    "Tripura": "",
    "Telangana": "",
    "Andhra Pradesh": "",
    "Karnataka": "",
    "Kerala": "",
    "Tamil Nadu": "",
    "Puducherry": "",
    "Delhi (NCT)": "",
    "Lakshadweep": "",
}

# =====================================================================
# 4. FIELD LABELS & ALIASES (100% Functionality Preserved)
# =====================================================================
FIELD_LABELS = {
    "Languages & Dialects": "Languages & Dialects",
    "Folk Dances": "Folk Dances",
    "Folk Music": "Folk Music",
    "Arts & Crafts": "Arts & Crafts",
    "Traditional Clothing": "Traditional Clothing",
    "Food & Cuisine": "Food & Cuisine",
    "Festivals & Celebrations": "Festivals & Celebrations",
    "Heritage & Monuments": "Heritage & Monuments",
    "History & Cultural Origin": "History & Cultural Origin",
    "Rituals & Traditions": "Rituals & Traditions",
    "Communities & Tribes": "Communities & Tribes",
    "Cultural Landmarks": "Cultural Landmarks",
    "Folk Tales & Legends": "Folk Tales & Legends",
    "Traditional Occupations": "Traditional Occupations",
    "Architecture": "Architecture",
    "Jewellery & Ornaments": "Jewellery & Ornaments",
    "Traditional Knowledge & Practices": "Traditional Knowledge & Practices",
    "Theatre & Performing Arts": "Theatre & Performing Arts",
    "Textiles & Weaving": "Textiles & Weaving",
}

ALIASES = {
    "jammu": "1.1 Jammu",
    "jammu and kashmir": "1.1 Jammu",
    "jammu & kashmir": "1.1 Jammu",
    "jammu kashmir": "1.1 Jammu",
    "j&k": "1.1 Jammu",
    "kashmir": "1.2 Kashmir",
    "kashmir valley": "1.2 Kashmir",
    "delhi": "Delhi (NCT)",
    "nct": "Delhi (NCT)",
    "new delhi": "Delhi (NCT)",
    "up": "Uttar Pradesh",
    "mp": "Madhya Pradesh",
    "hp": "Himachal Pradesh",
    "uk": "Uttarakhand",
    "ap": "Andhra Pradesh",
    "wb": "West Bengal",
    "bengal": "West Bengal",
    "orissa": "Odisha",
    "pondicherry": "Puducherry",
    "daman": "Dadra & Nagar Haveli and Daman & Diu",
    "diu": "Dadra & Nagar Haveli and Daman & Diu",
    "dadra": "Dadra & Nagar Haveli and Daman & Diu",
    "andaman": "Andaman & Nicobar Islands",
    "nicobar": "Andaman & Nicobar Islands",
}
# =====================================================================
# 4A. LOCATION INTELLIGENCE (Capital / City / District / Area)
# =====================================================================
LOCATION_ALIASES = {
    "bengaluru": "Karnataka", "bangalore": "Karnataka", "mysuru": "Karnataka",
    "mysore": "Karnataka", "hampi": "Karnataka", "udupi": "Karnataka",
    "coorg": "Karnataka", "kodagu": "Karnataka", "mangaluru": "Karnataka",
    "mangalore": "Karnataka", "hubballi": "Karnataka", "belagavi": "Karnataka",
    "thiruvananthapuram": "Kerala", "trivandrum": "Kerala", "kochi": "Kerala",
    "cochin": "Kerala", "kozhikode": "Kerala", "calicut": "Kerala",
    "thrissur": "Kerala", "alappuzha": "Kerala", "alleppey": "Kerala",
    "wayanad": "Kerala", "munnar": "Kerala",
    "chennai": "Tamil Nadu", "madurai": "Tamil Nadu", "thanjavur": "Tamil Nadu",
    "coimbatore": "Tamil Nadu", "ooty": "Tamil Nadu", "trichy": "Tamil Nadu",
    "rameswaram": "Tamil Nadu", "kanyakumari": "Tamil Nadu",
    "amaravati": "Andhra Pradesh", "tirupati": "Andhra Pradesh",
    "visakhapatnam": "Andhra Pradesh", "vizag": "Andhra Pradesh",
    "vijayawada": "Andhra Pradesh", "hyderabad": "Telangana",
    "warangal": "Telangana", "mumbai": "Maharashtra", "pune": "Maharashtra",
    "nagpur": "Maharashtra", "nashik": "Maharashtra", "kolhapur": "Maharashtra",
    "panaji": "Goa", "panjim": "Goa", "margao": "Goa",
    "ahmedabad": "Gujarat", "gandhinagar": "Gujarat", "vadodara": "Gujarat",
    "baroda": "Gujarat", "surat": "Gujarat", "kutch": "Gujarat", "bhuj": "Gujarat",
    "jaipur": "Rajasthan", "udaipur": "Rajasthan", "jodhpur": "Rajasthan",
    "jaisalmer": "Rajasthan", "pushkar": "Rajasthan", "bikaner": "Rajasthan",
    "lucknow": "Uttar Pradesh", "agra": "Uttar Pradesh", "varanasi": "Uttar Pradesh",
    "kashi": "Uttar Pradesh", "ayodhya": "Uttar Pradesh", "prayagraj": "Uttar Pradesh",
    "mathura": "Uttar Pradesh", "patna": "Bihar", "gaya": "Bihar",
    "bodh gaya": "Bihar", "ranchi": "Jharkhand", "deoghar": "Jharkhand",
    "bhopal": "Madhya Pradesh", "indore": "Madhya Pradesh", "ujjain": "Madhya Pradesh",
    "khajuraho": "Madhya Pradesh", "raipur": "Chhattisgarh", "jagdalpur": "Chhattisgarh",
    "amritsar": "Punjab", "ludhiana": "Punjab", "patiala": "Punjab",
    "chandigarh": "Chandigarh", "gurugram": "Haryana", "gurgaon": "Haryana",
    "kurukshetra": "Haryana", "shimla": "Himachal Pradesh", "manali": "Himachal Pradesh",
    "dharamshala": "Himachal Pradesh", "dehradun": "Uttarakhand", "haridwar": "Uttarakhand",
    "rishikesh": "Uttarakhand", "nainital": "Uttarakhand", "kedarnath": "Uttarakhand",
    "jammu": "1.1 Jammu", "srinagar": "1.2 Kashmir", "leh": "Ladakh",
    "gulmarg": "1.2 Kashmir", "pahalgam": "1.2 Kashmir",
    "kolkata": "West Bengal", "darjeeling": "West Bengal", "siliguri": "West Bengal",
    "bhubaneswar": "Odisha", "puri": "Odisha", "konark": "Odisha", "cuttack": "Odisha",
    "gangtok": "Sikkim", "guwahati": "Assam", "dispur": "Assam", "majuli": "Assam",
    "shillong": "Meghalaya", "cherrapunji": "Meghalaya", "sohra": "Meghalaya",
    "itanagar": "Arunachal Pradesh", "tawang": "Arunachal Pradesh",
    "kohima": "Nagaland", "dimapur": "Nagaland", "imphal": "Manipur",
    "aizawl": "Mizoram", "agartala": "Tripura",
    "delhi": "Delhi (NCT)", "new delhi": "Delhi (NCT)",
    "port blair": "Andaman & Nicobar Islands", "kavaratti": "Lakshadweep",
}

def find_location(text):
    low = text.lower()
    for place in sorted(LOCATION_ALIASES, key=len, reverse=True):
        if re.search(rf"\b{re.escape(place)}\b", low):
            return place.title(), LOCATION_ALIASES[place]
    return None, None

@st.cache_data(show_spinner=False, ttl=3600)
def geocode_location(place_query: str):
    if not place_query or len(place_query.strip()) < 3:
        return None
    try:
        res = requests.get(
            "https://nominatim.openstreetmap.org/search",
            params={
                "q": f"{place_query.strip()}, India",
                "format": "jsonv2",
                "limit": 1,
                "addressdetails": 1,
            },
            headers={"User-Agent": "CultureDNA-Chatbot/3.0"},
            timeout=5,
        )
        items = res.json()
        if not items:
            return None
        item = items[0]
        return {
            "place_name": item.get("display_name", place_query).split(",")[0],
            "lat": float(item["lat"]),
            "lon": float(item["lon"]),
            "state_name": item.get("address", {}).get("state", ""),
        }
    except Exception:
        return None

def state_from_geocoded_name(state_name):
    low = str(state_name).strip().lower()
    mapping = {
        "uttar pradesh": "Uttar Pradesh", "west bengal": "West Bengal",
        "tamil nadu": "Tamil Nadu", "andhra pradesh": "Andhra Pradesh",
        "telangana": "Telangana", "karnataka": "Karnataka", "kerala": "Kerala",
        "odisha": "Odisha", "orissa": "Odisha", "maharashtra": "Maharashtra",
        "rajasthan": "Rajasthan", "gujarat": "Gujarat", "punjab": "Punjab",
        "haryana": "Haryana", "bihar": "Bihar", "jharkhand": "Jharkhand",
        "chhattisgarh": "Chhattisgarh", "madhya pradesh": "Madhya Pradesh",
        "himachal pradesh": "Himachal Pradesh", "uttarakhand": "Uttarakhand",
        "assam": "Assam", "meghalaya": "Meghalaya", "sikkim": "Sikkim",
        "nagaland": "Nagaland", "manipur": "Manipur", "mizoram": "Mizoram",
        "tripura": "Tripura", "arunachal pradesh": "Arunachal Pradesh",
        "ladakh": "Ladakh", "goa": "Goa", "delhi": "Delhi (NCT)",
        "chandigarh": "Chandigarh", "puducherry": "Puducherry",
        "lakshadweep": "Lakshadweep",
        "andaman and nicobar islands": "Andaman & Nicobar Islands",
        "jammu and kashmir": "1.2 Kashmir",
    }
    return mapping.get(low)


# =====================================================================
# 5. STATE GPS COORDINATES (All 37 States/UTs in Culture_DNA.xlsx)
# =====================================================================
STATE_COORDINATES = {
    "1.1 Jammu": {
        "lat": 32.7266, "lon": 74.8570, "capital": "Jammu",
        "showcase": "Bahu Fort Jammu", "query": "Bahu Fort Jammu"
    },
    "1.2 Kashmir": {
        "lat": 34.0837, "lon": 74.7973, "capital": "Srinagar",
        "showcase": "Dal Lake Srinagar", "query": "Dal Lake Srinagar"
    },
    "Himachal Pradesh": {
        "lat": 31.1048, "lon": 77.1734, "capital": "Shimla",
        "showcase": "Shimla Ridge", "query": "The Ridge Shimla"
    },
    "Punjab": {
        "lat": 31.6200, "lon": 74.8765, "capital": "Amritsar / Chandigarh",
        "showcase": "Golden Temple Amritsar", "query": "Golden Temple Amritsar"
    },
    "Haryana": {
        "lat": 29.9695, "lon": 76.8783, "capital": "Chandigarh",
        "showcase": "Brahma Sarovar Kurukshetra", "query": "Brahma Sarovar Kurukshetra"
    },
    "Uttarakhand": {
        "lat": 30.7346, "lon": 79.0669, "capital": "Dehradun",
        "showcase": "Kedarnath Temple", "query": "Kedarnath Temple"
    },
    "Chandigarh": {
        "lat": 30.7525, "lon": 76.8101, "capital": "Chandigarh",
        "showcase": "Rock Garden of Chandigarh", "query": "Rock Garden of Chandigarh"
    },
    "Rajasthan": {
        "lat": 26.9239, "lon": 75.8267, "capital": "Jaipur",
        "showcase": "Hawa Mahal Jaipur", "query": "Hawa Mahal Jaipur"
    },
    "Gujarat": {
        "lat": 21.8380, "lon": 73.7191, "capital": "Gandhinagar",
        "showcase": "Statue of Unity", "query": "Statue of Unity"
    },
    "Maharashtra": {
        "lat": 18.9220, "lon": 72.8347, "capital": "Mumbai",
        "showcase": "Gateway of India Mumbai", "query": "Gateway of India Mumbai"
    },
    "Goa": {
        "lat": 15.5009, "lon": 73.9116, "capital": "Panaji",
        "showcase": "Basilica of Bom Jesus Goa", "query": "Basilica of Bom Jesus Goa"
    },
    "Madhya Pradesh": {
        "lat": 24.8318, "lon": 79.9199, "capital": "Bhopal",
        "showcase": "Khajuraho Group of Monuments", "query": "Khajuraho Group of Monuments"
    },
    "Dadra & Nagar Haveli and Daman & Diu": {
        "lat": 20.4162, "lon": 72.8344, "capital": "Daman",
        "showcase": "Moti Daman Fort", "query": "Moti Daman Fort"
    },
    "Uttar Pradesh": {
        "lat": 27.1751, "lon": 78.0421, "capital": "Lucknow",
        "showcase": "Taj Mahal Agra", "query": "Taj Mahal Agra"
    },
    "Bihar": {
        "lat": 24.6960, "lon": 84.9913, "capital": "Patna",
        "showcase": "Mahabodhi Temple Bodh Gaya", "query": "Mahabodhi Temple Bodh Gaya"
    },
    "Jharkhand": {
        "lat": 24.4925, "lon": 86.7001, "capital": "Ranchi",
        "showcase": "Baidyanath Temple Deoghar", "query": "Baidyanath Temple Deoghar"
    },
    "Chhattisgarh": {
        "lat": 19.2037, "lon": 81.7042, "capital": "Raipur",
        "showcase": "Chitrakote Falls", "query": "Chitrakote Falls"
    },
    "Odisha": {
        "lat": 19.8876, "lon": 86.0945, "capital": "Bhubaneswar",
        "showcase": "Konark Sun Temple", "query": "Konark Sun Temple"
    },
    "Ladakh": {
        "lat": 33.7595, "lon": 78.6674, "capital": "Leh",
        "showcase": "Pangong Tso Ladakh", "query": "Pangong Tso Ladakh"
    },
    "West Bengal": {
        "lat": 22.5448, "lon": 88.3426, "capital": "Kolkata",
        "showcase": "Victoria Memorial Kolkata", "query": "Victoria Memorial Kolkata"
    },
    "Sikkim": {
        "lat": 27.3297, "lon": 88.6122, "capital": "Gangtok",
        "showcase": "Rumtek Monastery", "query": "Rumtek Monastery"
    },
    "Assam": {
        "lat": 26.1664, "lon": 91.7054, "capital": "Dispur",
        "showcase": "Kamakhya Temple Guwahati", "query": "Kamakhya Temple Guwahati"
    },
    "Arunachal Pradesh": {
        "lat": 27.5861, "lon": 91.8656, "capital": "Itanagar",
        "showcase": "Tawang Monastery", "query": "Tawang Monastery"
    },
    "Meghalaya": {
        "lat": 25.2677, "lon": 91.7314, "capital": "Shillong",
        "showcase": "Living Root Bridges", "query": "Living Root Bridges Meghalaya"
    },
    "Andaman & Nicobar Islands": {
        "lat": 11.6739, "lon": 92.7483, "capital": "Port Blair",
        "showcase": "Cellular Jail Port Blair", "query": "Cellular Jail Port Blair"
    },
    "Nagaland": {
        "lat": 25.6747, "lon": 94.1105, "capital": "Kohima",
        "showcase": "Hornbill Festival Kohima", "query": "Hornbill Festival Nagaland"
    },
    "Manipur": {
        "lat": 24.5500, "lon": 93.8167, "capital": "Imphal",
        "showcase": "Loktak Lake Manipur", "query": "Loktak Lake Manipur"
    },
    "Mizoram": {
        "lat": 23.6933, "lon": 92.6074, "capital": "Aizawl",
        "showcase": "Reiek Heritage Village", "query": "Reiek Mizoram"
    },
    "Tripura": {
        "lat": 23.8344, "lon": 91.2825, "capital": "Agartala",
        "showcase": "Ujjayanta Palace Agartala", "query": "Ujjayanta Palace Agartala"
    },
    "Telangana": {
        "lat": 17.3616, "lon": 78.4747, "capital": "Hyderabad",
        "showcase": "Charminar Hyderabad", "query": "Charminar Hyderabad"
    },
    "Andhra Pradesh": {
        "lat": 13.6833, "lon": 79.3500, "capital": "Amaravati",
        "showcase": "Tirupati Balaji Temple", "query": "Venkateswara Temple Tirumala"
    },
    "Karnataka": {
        "lat": 15.3350, "lon": 76.4600, "capital": "Bengaluru",
        "showcase": "Hampi Monuments Karnataka", "query": "Hampi Monuments Karnataka"
    },
    "Kerala": {
        "lat": 9.4981, "lon": 76.3388, "capital": "Thiruvananthapuram",
        "showcase": "Kerala Backwaters & Kathakali", "query": "Kathakali dance Kerala"
    },
    "Tamil Nadu": {
        "lat": 9.9195, "lon": 78.1193, "capital": "Chennai",
        "showcase": "Meenakshi Amman Temple Madurai", "query": "Meenakshi Temple Madurai"
    },
    "Puducherry": {
        "lat": 12.0069, "lon": 79.8106, "capital": "Puducherry",
        "showcase": "Matrimandir Auroville", "query": "Matrimandir Auroville"
    },
    "Delhi (NCT)": {
        "lat": 28.6129, "lon": 77.2295, "capital": "New Delhi",
        "showcase": "India Gate New Delhi", "query": "India Gate New Delhi"
    },
    "Lakshadweep": {
        "lat": 10.9400, "lon": 72.2900, "capital": "Kavaratti",
        "showcase": "Bangaram Atoll Lakshadweep", "query": "Bangaram Atoll Lakshadweep"
    }
}

QUESTION_MAP = {
    "language": "Languages & Dialects",
    "languages": "Languages & Dialects",
    "dialect": "Languages & Dialects",
    "dance": "Folk Dances",
    "dances": "Folk Dances",
    "folk dance": "Folk Dances",
    "music": "Folk Music",
    "folk music": "Folk Music",
    "craft": "Arts & Crafts",
    "crafts": "Arts & Crafts",
    "art": "Arts & Crafts",
    "dress": "Traditional Clothing",
    "clothing": "Traditional Clothing",
    "attire": "Traditional Clothing",
    "food": "Food & Cuisine",
    "foods": "Food & Cuisine",
    "cuisine": "Food & Cuisine",
    "dish": "Food & Cuisine",
    "dishes": "Food & Cuisine",
    "festival": "Festivals & Celebrations",
    "festivals": "Festivals & Celebrations",
    "celebration": "Festivals & Celebrations",
    "heritage": "Heritage & Monuments",
    "monument": "Heritage & Monuments",
    "monuments": "Heritage & Monuments",
    "history": "History & Cultural Origin",
    "origin": "History & Cultural Origin",
    "tradition": "Rituals & Traditions",
    "traditions": "Rituals & Traditions",
    "ritual": "Rituals & Traditions",
    "community": "Communities & Tribes",
    "communities": "Communities & Tribes",
    "tribe": "Communities & Tribes",
    "tribes": "Communities & Tribes",
    "landmark": "Cultural Landmarks",
    "landmarks": "Cultural Landmarks",
    "legend": "Folk Tales & Legends",
    "legends": "Folk Tales & Legends",
    "occupation": "Traditional Occupations",
    "occupations": "Traditional Occupations",
    "architecture": "Architecture",
    "jewellery": "Jewellery & Ornaments",
    "jewelry": "Jewellery & Ornaments",
    "ornaments": "Jewellery & Ornaments",
    "knowledge": "Traditional Knowledge & Practices",
    "practices": "Traditional Knowledge & Practices",
    "theatre": "Theatre & Performing Arts",
    "theater": "Theatre & Performing Arts",
    "performing arts": "Theatre & Performing Arts",
    "textile": "Textiles & Weaving",
    "textiles": "Textiles & Weaving",
    "weaving": "Textiles & Weaving",
}

def list_states():
    if not df.empty and "State/UT" in df.columns:
        return [str(x) for x in df["State/UT"].tolist()]
    return list(STATE_COORDINATES.keys())

def clean_state_name(name):
    name = str(name).strip()
    states_all = list_states()
    if name in states_all:
        return name
    low = name.lower()
    if low in ALIASES:
        return ALIASES[low]
    for state in states_all:
        if low == str(state).lower():
            return state
    return None

def find_state(text):
    text_low = text.lower()
    if "kashmir valley" in text_low or re.search(r"\bkashmir\b", text_low):
        if "jammu" not in text_low:
            return "1.2 Kashmir"
    if "jammu" in text_low or "j&k" in text_low:
        return "1.1 Jammu"

    for state in list_states():
        state_low = str(state).lower()
        if state_low in text_low:
            return state

    for alias, state in ALIASES.items():
        if len(alias) <= 3:
            if re.search(rf"\b{re.escape(alias)}\b", text_low):
                return state
        elif alias in text_low:
            return state

    return None

def find_field(text):
    """Resolve a cultural category from QUESTION_MAP *or directly from Excel columns*.
    This keeps image/video galleries dynamic for every cultural question/column in the
    workbook instead of limiting the bot to a hard-coded list of categories.
    """
    text_low = str(text).lower()

    # 1) Friendly/common words first.
    for key, field in sorted(QUESTION_MAP.items(), key=lambda x: len(x[0]), reverse=True):
        if re.search(rf"\b{re.escape(key)}\b", text_low):
            return field

    # 2) Automatically recognize every non-system column present in Culture_DNA.xlsx.
    #    Example: if Excel later gets a new column called "Traditional Games",
    #    a query such as "games of Bihar" can resolve to that column.
    if not df.empty:
        ignored = {"state/ut", "image_url", "multimedia / images", "multimedia/images"}
        candidates = []
        for col in df.columns:
            col_text = str(col).strip()
            if not col_text or col_text.lower() in ignored:
                continue
            # Exact phrase match is strongest.
            norm_col = re.sub(r"[^a-z0-9]+", " ", col_text.lower()).strip()
            if norm_col and norm_col in text_low:
                candidates.append((len(norm_col), col_text))
                continue
            # Otherwise require useful content words from the Excel header.
            words = [w for w in re.findall(r"[a-z0-9]+", norm_col) if len(w) >= 4]
            if words:
                hits = sum(1 for w in words if re.search(rf"\b{re.escape(w)}\b", text_low))
                if hits >= max(1, len(words) // 2):
                    candidates.append((hits * 100 + len(norm_col), col_text))
        if candidates:
            candidates.sort(reverse=True)
            return candidates[0][1]

    return None

def row_for_state(state):
    if df.empty or "State/UT" not in df.columns:
        return None
    rows = df[df["State/UT"].astype(str) == state]
    return rows.iloc[0] if not rows.empty else None

def get_clean_region_name(state: str) -> str:
    s = str(state).strip()
    if s.startswith("1.1"):
        return "Jammu"
    if s.startswith("1.2"):
        return "Kashmir"
    return s

def get_display_name(state: str) -> str:
    s = str(state).strip()
    if s.startswith("1.1"):
        return "Jammu Region"
    if s.startswith("1.2"):
        return "Kashmir Valley"
    return s

# =====================================================================
# 6. CULTURAL VISUAL & GEO ENGINE
# =====================================================================
@st.cache_data(show_spinner=False, ttl=3600)
def fetch_cultural_visual(search_query: str):
    """
    Queries Wikipedia API for authentic high-resolution images and coordinates.
    Cached for fast response.
    """
    if not search_query or not search_query.strip():
        return None

    api_url = "https://en.wikipedia.org/w/api.php"
    params = {
        "action": "query",
        "format": "json",
        "generator": "search",
        "gsrsearch": search_query.strip(),
        "gsrlimit": 1,
        "prop": "pageimages|coordinates",
        "pithumbsize": 800
    }
    headers = {"User-Agent": "CultureDNA-Chatbot/2.0 (visual-cultural-guide)"}

    try:
        res = requests.get(api_url, params=params, headers=headers, timeout=4).json()
        pages = res.get("query", {}).get("pages", {})
        for _, data in pages.items():
            img_url = None
            if "thumbnail" in data:
                img_url = data["thumbnail"]["source"]

            coords = None
            if "coordinates" in data and len(data["coordinates"]) > 0:
                coords = {
                    "lat": float(data["coordinates"][0]["lat"]),
                    "lon": float(data["coordinates"][0]["lon"])
                }

            page_title = data.get("title", search_query)
            if img_url or coords:
                return {
                    "image_url": img_url,
                    "coords": coords,
                    "title": page_title
                }
    except Exception:
        pass

    return None

def youtube_search_url(query: str):
    return "https://www.youtube.com/results?search_query=" + quote_plus(query) if query else None

def fetch_commons_topic_image(search_query: str):
    """Find a Wikimedia Commons image that is genuinely about the requested item.
    Generic city/building/landmark images are rejected for topic questions.
    """
    if not search_query or not search_query.strip():
        return None

    api_url = "https://commons.wikimedia.org/w/api.php"
    params = {
        "action": "query", "format": "json", "generator": "search",
        "gsrsearch": search_query.strip(), "gsrnamespace": 6, "gsrlimit": 12,
        "prop": "imageinfo", "iiprop": "url|extmetadata", "iilimit": 1,
    }
    headers = {"User-Agent": "CultureDNA-Chatbot/4.0 (topic-image-search)"}

    # Words that usually indicate a generic place photo rather than the requested
    # food/dress/dance/craft/etc.
    bad_words = {
        "landmark", "building", "monument", "palace", "station", "temple",
        "church", "mosque", "fort", "street", "city", "skyline", "aerial",
        "view", "panorama", "map", "district", "railway", "airport"
    }

    try:
        res = requests.get(api_url, params=params, headers=headers, timeout=7).json()
        pages = res.get("query", {}).get("pages", {})
        # Score ONLY the cultural item terms, not helper words such as the
        # state name or "India".  Otherwise valid files like "Sattu.jpg"
        # get rejected simply because their title does not also contain
        # "Bihar" and "India".
        query_words = [w for w in (required_terms or re.findall(r"[a-z0-9]+", search_query.lower())) if len(w) >= 3]
        candidates = []

        for _, data in pages.items():
            info = data.get("imageinfo", [{}])[0]
            url = info.get("thumburl") or info.get("url")
            title = data.get("title", "")
            if not url:
                continue
            title_low = title.lower()
            score = 0

            # Strong match: all meaningful query words appear in the file title.
            matched = sum(1 for w in query_words if w in title_low)
            score += matched * 5
            if query_words and matched == len(query_words):
                score += 15

            if any(b in title_low for b in bad_words):
                score -= 20

            # Never accept a generic place image when the requested item has
            # multiple meaningful words and none of them appears in the title.
            if len(query_words) >= 2 and matched < 2:
                continue

            candidates.append((score, title, url))

        if candidates:
            candidates.sort(key=lambda x: x[0], reverse=True)
            score, title, url = candidates[0]
            if score >= 10:
                return {"image_url": url, "title": title.replace("File:", "").strip()}
    except Exception:
        pass
    return None


@st.cache_data(show_spinner=False, ttl=3600)
def fetch_location_wikipedia(query: str):
    """Fetch location-specific cultural context from Wikipedia without replacing the dataset."""
    if not query or len(query.strip()) < 3:
        return None
    try:
        api_url = "https://en.wikipedia.org/w/api.php"
        params = {
            "action": "query", "format": "json", "generator": "search",
            "gsrsearch": query.strip(), "gsrlimit": 3,
            "prop": "extracts|pageimages|coordinates",
            "exintro": 1, "explaintext": 1, "exchars": 1200,
            "pithumbsize": 900
        }
        headers = {"User-Agent": "CultureDNA-Chatbot/4.0 (location-cultural-guide)"}
        res = requests.get(api_url, params=params, headers=headers, timeout=6).json()
        pages = res.get("query", {}).get("pages", {})
        if not pages:
            return None
        ranked = []
        for data in pages.values():
            title = str(data.get("title", ""))
            extract = str(data.get("extract", "")).strip()
            score = 0
            qwords = [w for w in re.findall(r"[a-zA-Z]{4,}", query.lower())]
            title_low = title.lower()
            for w in qwords:
                if w in title_low:
                    score += 3
            if extract:
                score += 2
            ranked.append((score, data))
        ranked.sort(key=lambda x: x[0], reverse=True)
        data = ranked[0][1]
        image_url = data.get("thumbnail", {}).get("source")
        coords = None
        if data.get("coordinates"):
            coords = {"lat": float(data["coordinates"][0]["lat"]), "lon": float(data["coordinates"][0]["lon"])}
        return {"title": data.get("title", query), "extract": extract, "image_url": image_url, "coords": coords}
    except Exception:
        return None

def district_culture_reply(location: str, state: str, field: str, user_query: str):
    """Build a district/city-specific answer when a place is more specific than the state."""
    label = FIELD_LABELS.get(field, field) if field else "Cultural Highlights"
    topic_terms = {
        "Food & Cuisine": "food cuisine famous dishes",
        "Traditional Clothing": "traditional dress clothing attire",
        "Folk Dances": "folk dance traditional dance",
        "Folk Music": "folk music traditional music",
        "Festivals & Celebrations": "festivals celebrations",
        "Arts & Crafts": "arts crafts handicrafts",
        "Heritage & Monuments": "famous places monuments heritage",
        "Cultural Landmarks": "famous places cultural landmarks",
        "Architecture": "traditional architecture famous buildings",
        "Jewellery & Ornaments": "traditional jewellery ornaments",
        "Textiles & Weaving": "traditional textiles weaving",
        "Theatre & Performing Arts": "traditional theatre performing arts",
        "Folk Tales & Legends": "folk tales legends",
        "Rituals & Traditions": "traditions rituals",
        "Traditional Occupations": "traditional occupations",
        "Communities & Tribes": "communities tribes culture",
        "Languages & Dialects": "languages dialects",
        "History & Cultural Origin": "history culture origin",
        "Traditional Knowledge & Practices": "traditional knowledge practices",
    }
    topic = topic_terms.get(field, "culture cultural heritage")
    wiki = fetch_location_wikipedia(f"{location} {topic} {state}")
    if not wiki:
        wiki = fetch_location_wikipedia(f"{location} {state} culture")
    if wiki:
        text = wiki.get("extract", "")
        if text:
            text = text[:1200].rstrip()
        reply = (f"### 📍 {location} — {label}\n\n"
                 f"**State/UT:** {get_display_name(state)}\n\n"
                 f"{text if text else 'I found a location-specific cultural reference, but no short description was available.'}")
        return reply, wiki
    return (f"### 📍 {location} — {label}\n\n"
            f"I could identify **{location}** in **{get_display_name(state)}**, but I could not find a reliable location-specific reference right now. "
            "I will not substitute generic state information for this district/city question."), None


def _split_cultural_items(value: str):
    """Return all meaningful cultural subjects stored in a dataset cell."""
    items = []
    for part in re.split(r"[,;\n]+", str(value or "")):
        part = re.sub(r"\([^)]*\)", "", part).strip()
        part = re.sub(r"^(include|including|such as|major|famous|popular)\s+", "", part, flags=re.IGNORECASE).strip(" -:")
        if len(part) > 2 and part.lower() not in {x.lower() for x in items}:
            items.append(part)
    return items


def fetch_commons_topic_images(search_query: str, max_images: int = 2, required_terms=None):
    """Find several exact-topic Wikimedia Commons images; reject generic place photos."""
    if not search_query or not search_query.strip():
        return []
    api_url = "https://commons.wikimedia.org/w/api.php"
    params = {
        "action": "query", "format": "json", "generator": "search",
        "gsrsearch": search_query.strip(), "gsrnamespace": 6, "gsrlimit": 30,
        "prop": "imageinfo", "iiprop": "url|extmetadata", "iilimit": 1,
    }
    headers = {"User-Agent": "CultureDNA-Chatbot/5.0 (exact-topic-gallery)"}
    bad_words = {
        "landmark", "building", "monument", "palace", "station", "temple",
        "church", "mosque", "fort", "street", "city", "skyline", "aerial",
        "view", "panorama", "map", "district", "railway", "airport"
    }
    try:
        res = requests.get(api_url, params=params, headers=headers, timeout=8).json()
        pages = res.get("query", {}).get("pages", {})
        # Score ONLY the cultural item terms, not helper words such as the
        # state name or "India".  Otherwise valid files like "Sattu.jpg"
        # get rejected simply because their title does not also contain
        # "Bihar" and "India".
        query_words = [w for w in (required_terms or re.findall(r"[a-z0-9]+", search_query.lower())) if len(w) >= 3]
        candidates = []
        for data in pages.values():
            info = data.get("imageinfo", [{}])[0]
            url = info.get("thumburl") or info.get("url")
            title = data.get("title", "")
            if not url:
                continue
            title_low = title.lower()
            matched = sum(1 for w in query_words if w in title_low)
            score = matched * 5
            if query_words and matched == len(query_words):
                score += 20
            if any(b in title_low for b in bad_words):
                score -= 25
            if len(query_words) >= 2 and matched < 2:
                continue
            if score < 10:
                continue
            candidates.append((score, title.replace("File:", "").strip(), url))
        candidates.sort(key=lambda x: x[0], reverse=True)
        out, seen = [], set()
        for _, title, url in candidates:
            if url in seen:
                continue
            seen.add(url)
            out.append({"image_url": url, "title": title})
            if len(out) >= max_images:
                break
        return out
    except Exception:
        return []


def get_visual_for_reply(state: str, field: str = None, row=None, location: str = None, location_coords=None, user_query: str = ""):
    """Return a gallery of exact cultural subjects, not only the first item.
    Example: Food of Bihar -> images for every food item present in the dataset.
    """
    state_meta = STATE_COORDINATES.get(state, {})
    clean_name = get_clean_region_name(state)
    display_title = get_display_name(state)
    location_label = location.strip() if location else display_title
    topic_label = FIELD_LABELS.get(field, field) if field else None

    # IMPORTANT: collect ALL items from the dataset cell.
    topic_items = []
    if field and row is not None:
        topic_items = _split_cultural_items(row.get(field, ""))

    gallery = []
    visual_result = None
    if topic_items:
        for exact_item in topic_items:
            item_clean = exact_item.replace("-", " ").strip()
            item_found = None
            exact_queries = [
                f'"{item_clean}" {clean_name} India',
                f'{item_clean} {clean_name} India',
                f'{item_clean} India',
            ]
            for q in exact_queries:
                imgs = fetch_commons_topic_images(q, max_images=2, required_terms=[w for w in re.findall(r"[a-z0-9]+", item_clean.lower()) if len(w) >= 3])
                if imgs:
                    item_found = imgs
                    break
            # Wikipedia fallback: at most one exact image for this item.
            if not item_found:
                for q in exact_queries[:2]:
                    result = fetch_cultural_visual(q)
                    if result and result.get("image_url"):
                        title_low = str(result.get("title", "")).lower()
                        required = [w for w in re.findall(r"[a-z0-9]+", item_clean.lower()) if len(w) >= 3]
                        if all(w in title_low for w in required):
                            item_found = [{"image_url": result["image_url"], "title": result.get("title", exact_item)}]
                            break
            if item_found:
                for img in item_found:
                    img["topic_item"] = exact_item
                    gallery.append(img)
            # Never replace a missing exact item with a random generic image.

    # If the user explicitly named a subject that isn't represented in the cell,
    # search that subject exactly as a secondary path.
    if not gallery and field:
        low_query = user_query.lower()
        topic_words = {
            "food": "food", "dance": "folk dance", "dress": "traditional clothing",
            "clothing": "traditional clothing", "attire": "traditional clothing",
            "music": "folk music", "festival": "festival", "craft": "traditional craft",
            "monument": "monument", "heritage": "heritage"
        }
        subject = next((v for k, v in topic_words.items() if re.search(rf"\b{k}\b", low_query)), topic_label or "culture")
        for q in [f"{subject} {location_label} {clean_name} India", f"{subject} {clean_name} India"]:
            imgs = fetch_commons_topic_images(q, max_images=3)
            if imgs:
                gallery = imgs
                break

    # Generic location/state imagery is ONLY allowed when no topic was requested.
    custom_image_url = None
    if not field and row is not None and "image_url" in row:
        val = str(row["image_url"]).strip()
        if val.startswith(("http://", "https://")):
            custom_image_url = val
    if not field and not custom_image_url and state in STATE_CUSTOM_IMAGE_URLS:
        custom_image_url = STATE_CUSTOM_IMAGE_URLS[state].strip()

    if not field and not gallery and location:
        visual_result = fetch_cultural_visual(f"{location} {clean_name} India") or fetch_cultural_visual(location)
    if not field and not gallery and state in STATE_COORDINATES:
        visual_result = fetch_cultural_visual(state_meta.get("query", f"{clean_name} landmark"))
    if not field and not gallery:
        visual_result = fetch_cultural_visual(f"{clean_name} India")

    if location_coords:
        final_coords = {"lat": location_coords["lat"], "lon": location_coords["lon"]}
        location_name = location_label
    elif visual_result and visual_result.get("coords"):
        final_coords = visual_result["coords"]
        location_name = visual_result.get("title", location_label)
    elif state_meta:
        final_coords = {"lat": state_meta["lat"], "lon": state_meta["lon"]}
        location_name = f"{display_title} ({state_meta.get('capital', 'Region')})"
    else:
        final_coords = None
        location_name = location_label

    # Exact item gets an exact video search for EACH item.
    video_items = topic_items if topic_items else ([gallery[0].get("topic_item")] if gallery and gallery[0].get("topic_item") else [])
    videos = []
    for item in video_items:
        if item:
            videos.append({"item": item, "url": youtube_search_url(f"{item} {clean_name} India")})
    if not videos:
        if topic_label:
            q = f"{location_label} {clean_name} {topic_label} India"
        else:
            q = f"{location_label} {clean_name} India culture"
        videos = [{"item": topic_label or location_label, "url": youtube_search_url(q)}]

    return {
        "image_url": custom_image_url or (gallery[0].get("image_url") if gallery else (visual_result.get("image_url") if visual_result else None)),
        "image_caption": gallery[0].get("title") if gallery else (visual_result.get("title") if visual_result else (topic_items[0] if topic_items else topic_label or location_label)),
        "images": gallery,
        "topic_items": topic_items,
        "coords": final_coords,
        "location_name": location_name,
        "video_url": videos[0]["url"] if videos else None,
        "video_query": f"{videos[0]['item']} {clean_name} India" if videos else "",
        "videos": videos,
        "topic_item": topic_items[0] if topic_items else None,
    }


def render_visual_gallery(visual, heading="📸 Exact Cultural Images"):
    """Render every exact cultural item image returned by the visual engine."""
    images = visual.get("images") or []
    if images:
        st.markdown(f"**{heading}**")
        cols = st.columns(2 if len(images) > 1 else 1)
        for i, img in enumerate(images):
            with cols[i % len(cols)]:
                st.image(img["image_url"], caption=f"{img.get('topic_item', 'Cultural subject')} — {img.get('title', '')}".strip(" —"), use_container_width=True)
        if visual.get("topic_items"):
            found = {str(x.get("topic_item", "")).lower() for x in images}
            missing = [x for x in visual["topic_items"] if x.lower() not in found]
            if missing:
                st.caption("Exact image not found for: " + ", ".join(missing) + ". No generic image was substituted.")
    elif visual.get("image_url"):
        st.markdown(f"**{heading}**")
        st.image(visual["image_url"], caption=visual.get("image_caption", "Cultural Heritage"), use_container_width=True)
    elif visual.get("topic_items"):
        st.info("No exact images were found for the listed cultural subjects yet; no random images were shown.")

    videos = visual.get("videos") or []
    if videos:
        st.markdown("**🎥 Exact Cultural Videos**")
        for video in videos:
            st.link_button(f"▶️ {video.get('item', 'Relevant cultural video')}", video["url"], use_container_width=True)

# =====================================================================
# 7. INTERACTIVE FOLIUM MAP (Requirement 3 & 8)
# Automatically centers and moves to [lat, lon] using dynamic key
# =====================================================================
def render_interactive_map(lat: float, lon: float, title: str, map_key: str = None, zoom: int = 7):
    """
    Renders an interactive Folium map centered on [lat, lon] with location pin.
    Automatically moves to the selected state using dynamic key.
    """
    m = folium.Map(
        location=[lat, lon],
        zoom_start=zoom,
        tiles="OpenStreetMap",
        control_scale=True
    )
    folium.Marker(
        [lat, lon],
        popup=folium.Popup(f"<b>{title}</b>", max_width=300),
        tooltip=f"📍 {title}",
        icon=folium.Icon(color="red", icon="info-sign")
    ).add_to(m)

    k = map_key or f"map_{lat}_{lon}"
    st_folium(
        m,
        height=320,
        use_container_width=True,
        returned_objects=[],
        key=k
    )

# =====================================================================
# 7A. VOICE INPUT
# =====================================================================
def transcribe_voice(audio_data, language_code="en-IN"):
    """Convert microphone WAV bytes to text using Google Speech Recognition."""
    if not audio_data:
        return ""
    try:
        recognizer = sr.Recognizer()
        audio_bytes = audio_data.get("bytes") if isinstance(audio_data, dict) else audio_data
        if not audio_bytes:
            return ""
        with sr.AudioFile(__import__("io").BytesIO(audio_bytes)) as source:
            audio = recognizer.record(source)
        return recognizer.recognize_google(audio, language=language_code)
    except sr.UnknownValueError:
        return ""
    except sr.RequestError:
        return "__VOICE_SERVICE_ERROR__"
    except Exception:
        return ""

# =====================================================================
# 8. CHATBOT CORE LOGIC
# =====================================================================
def chatbot_reply(message: str, selected_state: str = None):
    text = message.strip()
    low = text.lower()

    if not text:
        return {"text": "Please ask me about a state, city, capital, region, or cultural place. 😊", "visual": None}

    if re.search(r"\b(hello|hi|hey|namaste|pranam)\b", low):
        return {
            "text": (
                "Namaste! 👋 I am **Culture DNA**, your visual & cultural guide to India.\n\n"
                "Ask me about a **state, capital, city, district, region, monument, food, "
                "folk dance, festival, attire, music, crafts, rituals, or history**.\n\n"
                "Try: *'Culture of Bengaluru'*, *'Food of Mysuru'*, or *'What is famous in Hampi?'*"
            ),
            "visual": None
        }

    # Functional actions coming from the visual home screen.
    if any(x in low for x in ["explore india on map", "show india map", "map of india"]):
        return {
            "text": (
                "### 🗺️ Explore India on Map\n\n"
                "The interactive India map is shown below. You can zoom, pan and inspect the map. "
                "For a specific location, try **'Map of Patna'** or **'Where is Mysuru?'**."
            ),
            "visual": {
                "image_url": None,
                "image_caption": "India Map",
                "coords": {"lat": 22.9734, "lon": 78.6569},
                "location_name": "India",
                "video_url": None,
                "video_query": "",
                "topic_item": None
            }
        }

    if "what is culture dna" in low or "tell me about culture dna" in low:
        return {
            "text": (
                "### 🧬 About Culture DNA\n\n"
                "Culture DNA is a visual cultural guide for India. It helps you explore "
                "state, district and city-level culture through structured information, "
                "authentic cultural visuals, relevant videos and interactive maps.\n\n"
                "You can ask naturally in English or Hindi/Hinglish."
            ),
            "visual": None
        }

    if "state-wise cultural information" in low:
        return {
            "text": (
                "### 📖 State-wise Cultural Information\n\n"
                "Select a State/UT from the sidebar or ask directly, for example **'Culture of Bihar'**, "
                "**'Food of Karnataka'**, or **'Dance of Assam'**. The answer can include an exact cultural "
                "visual, relevant video and an interactive location map."
            ),
            "visual": None
        }

    if "authentic cultural images" in low:
        return {
            "text": (
                "### 🖼️ Authentic Cultural Images\n\n"
                "Images are matched to the cultural subject you ask for. For example, asking "
                "**'Food of Bihar'** searches for the exact food item rather than a random Indian food image."
            ),
            "visual": None
        }

    if "relevant cultural videos" in low:
        return {
            "text": (
                "### ▶️ Relevant Cultural Videos\n\n"
                "Video searches are built from the exact cultural subject and location, so a food, dress, "
                "dance or festival request gets a matching video search instead of a generic culture video."
            ),
            "visual": None
        }

    if "which states" in low or "list states" in low or "all states" in low:
        return {
            "text": "I currently have cultural records for these States & UTs:\n\n" + "\n".join(
                f"- {s}" for s in list_states()
            ),
            "visual": None
        }

    state = find_state(text)
    location, location_state = find_location(text)

    if not state and location_state:
        state = location_state

    geocoded = None
    if not state:
        query_for_geo = re.sub(
            r"\b(what|which|tell|me|about|the|culture|of|in|from|is|are|"
            r"famous|traditional|food|dance|festival|heritage|monument|"
            r"district|city|town|village|area|place|show|where|map|please)\b",
            " ", low, flags=re.IGNORECASE
        )
        query_for_geo = re.sub(r"\s+", " ", query_for_geo).strip()
        if query_for_geo and len(query_for_geo) >= 3:
            geocoded = geocode_location(query_for_geo)
            if geocoded:
                state = state_from_geocoded_name(geocoded.get("state_name", ""))
                location = geocoded.get("place_name") or query_for_geo

    if not state:
        state = selected_state

    field = find_field(text)

    if state:
        row = row_for_state(state)
        if row is None:
            return {
                "text": "I found the location, but it is not present in the current Culture DNA dataset.",
                "visual": None
            }

        display_name = get_display_name(state)
        location_coords = (
            {"lat": geocoded["lat"], "lon": geocoded["lon"]}
            if geocoded else None
        )

        # If a city/district/place is named, do NOT answer with the whole state's
        # generic row. Use a location-specific source instead.
        state_like_location = {display_name.lower(), get_clean_region_name(state).lower(), str(state).lower()}
        is_specific_location = bool(location and location.strip().lower() not in state_like_location)
        if is_specific_location:
            district_text, district_ref = district_culture_reply(location, state, field, text)
            topic_for_visual = field
            visual_data = get_visual_for_reply(
                state, topic_for_visual, row, location=location, location_coords=location_coords, user_query=text
            )
            # Prefer the location-specific Wikipedia image when it exists; for a topic,
            # get_visual_for_reply still enforces the exact dataset item when available.
            if district_ref and district_ref.get("image_url") and not field:
                visual_data["image_url"] = district_ref["image_url"]
                visual_data["image_caption"] = district_ref.get("title", location)
            if district_ref and district_ref.get("coords") and not location_coords:
                visual_data["coords"] = district_ref["coords"]
                visual_data["location_name"] = district_ref.get("title", location)
            if field:
                # Keep exact dataset item visible as a related state reference, but never
                # pretend it is district-specific if the external source could not verify it.
                district_text += f"\n\n**Requested category:** {FIELD_LABELS.get(field, field)}"
            return {"text": district_text, "visual": visual_data}

        visual_data = get_visual_for_reply(
            state, field, row, location=location, location_coords=location_coords, user_query=text
        )

        context_line = (
            f"**📍 Location:** {location} · **State/UT:** {display_name}\n\n"
            if location else ""
        )

        if any(w in low for w in ["map", "location", "locate", "kahan h", "kahan hai", "where is"]):
            if location_coords:
                reply_text = (
                    f"### 🗺️ {location} — Geographic Location\n\n"
                    f"- **State/UT:** {display_name}\n"
                    f"- **Coordinates:** `{location_coords['lat']:.4f}° N, {location_coords['lon']:.4f}° E`\n\n"
                    f"The interactive map below is centered on **{location}**."
                )
            else:
                meta = STATE_COORDINATES.get(state, {})
                reply_text = (
                    f"### 🗺️ {display_name} — Geographic Location\n\n"
                    f"- **Region / Capital:** {meta.get('capital', 'Region')}\n"
                    f"- **Coordinates:** `{meta.get('lat', 0):.2f}° N, {meta.get('lon', 0):.2f}° E`\n\n"
                    "Interactive Folium map is displayed below."
                )
            return {"text": reply_text, "visual": visual_data}

        if field:
            value = str(row.get(field, "")).strip()
            if not value:
                reply_text = (
                    f"{context_line}I don't have verified data for **{FIELD_LABELS[field]}** "
                    f"for {display_name} yet."
                )
            else:
                reply_text = (
                    f"### 🏛️ {location or display_name}\n\n"
                    f"{context_line}**{FIELD_LABELS[field]}:**\n{value}"
                )
            return {"text": reply_text, "visual": visual_data}

        # General profile: include ALL cultural columns present in Excel, not just
        # the original eight categories. This makes the chatbot future-proof when
        # additional cultural questions are added to the workbook.
        parts = []
        system_cols = {"state/ut", "image_url", "multimedia / images", "multimedia/images"}
        for col in df.columns if not df.empty else []:
            col_name = str(col).strip()
            if col_name.lower() in system_cols or not col_name:
                continue
            value = str(row.get(col, "")).strip()
            if value:
                parts.append(f"**{FIELD_LABELS.get(col, col_name)}:** {value}")

        title = location or display_name
        reply_text = f"## 🧬 Culture DNA — {title}\n\n{context_line}" + "\n\n".join(parts)
        return {"text": reply_text, "visual": visual_data}

    if field:
        return {
            "text": (
                f"I can answer about **{FIELD_LABELS[field]}**. Include a state, capital, city, "
                "or region, e.g. **'food of Bengaluru'**, **'festivals of Mysuru'**, or **'dance of Assam'**."
            ),
            "visual": None
        }

    return {
        "text": (
            "I can help you explore **states, capitals, cities and cultural places**.\n\n"
            "Try: **'Culture of Bengaluru'**, **'Food of Mysuru'**, or **'What is famous in Hampi?'**"
        ),
        "visual": None
    }

# =====================================================================
# HOME FRONTEND ACTIONS — every visible card/arrow is functional
def consume_home_action():
    try:
        prompt = st.query_params.get("home_prompt")
        action = st.query_params.get("home_action")
        if prompt:
            st.session_state["queued_prompt"] = prompt
            st.query_params.clear()
        elif action == "map":
            st.session_state["queued_prompt"] = "Explore India on map"
            st.query_params.clear()
        elif action == "chat":
            st.session_state["queued_prompt"] = "Tell me about Culture DNA"
            st.query_params.clear()
        elif action == "about":
            st.session_state["queued_prompt"] = "What is Culture DNA?"
            st.query_params.clear()
    except Exception:
        pass

consume_home_action()

# 9. UI LAYOUT & DISPLAY — REAL FRONTEND HOME (not a screenshot)
# =====================================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Playfair+Display:wght@600;700;800&family=Caveat:wght@500;600&display=swap');

/* App background */
.stApp {
    background: linear-gradient(180deg, #071f33 0%, #0a2b43 16%, #f8ead8 16%, #fff8ef 100%);
}
.block-container { max-width: 1400px; padding-top: 1rem; }

.home-front {
    overflow: hidden;
    border-radius: 0 0 34px 34px;
    background: #fff7eb;
    box-shadow: 0 20px 55px rgba(7,31,51,.22);
}
.home-nav {
    min-height: 92px;
    padding: 18px 34px;
    background: linear-gradient(100deg,#06223a,#0b3550);
    color: white;
    display:flex;
    align-items:center;
    justify-content:space-between;
    gap:25px;
}
.brand {display:flex;align-items:center;gap:14px;}
.brand-flag {font-size:2.4rem;line-height:1;}
.brand-name {font:800 1.7rem Inter,sans-serif;letter-spacing:.5px;}
.brand-name span {color:#f47a28;}
.brand-sub {font:500 .72rem Inter,sans-serif;letter-spacing:4px;opacity:.92;margin-top:3px;}
.nav-links {display:flex;gap:30px;font:600 .92rem Inter,sans-serif;align-items:center;}
.nav-links span {opacity:.95;} .nav-links a{color:inherit;text-decoration:none;opacity:.95}.nav-links a:hover{color:#ffd3a8}
.motto {font:italic 1.25rem 'Playfair Display',serif;line-height:1.15;text-align:right;}
.motto i {display:inline-block;width:42px;border-bottom:3px solid #f28b35;margin-top:10px;}

.hero {
    position:relative;
    padding:55px 54px 28px;
    min-height:520px;
    background: radial-gradient(circle at 78% 28%, rgba(252,195,116,.55), transparent 30%), linear-gradient(120deg,#fff8ee,#f8ead8 55%,#e9d9c4);
    overflow:hidden;
}
.hero:before {content:"";position:absolute;left:-100px;top:45px;width:500px;height:120px;border-top:28px solid rgba(245,126,37,.95);border-radius:50%;transform:rotate(-10deg);}
.hero:after {content:"";position:absolute;left:-100px;top:95px;width:480px;height:100px;border-top:18px solid rgba(33,121,83,.82);border-radius:50%;transform:rotate(-10deg);}
.hero-copy {position:relative;z-index:2;max-width:650px;}
.kicker {display:inline-block;padding:9px 18px;border-radius:30px;background:#fff0dc;color:#7b4a28;font:700 .82rem Inter,sans-serif;letter-spacing:1.4px;}
.hero h1 {margin:22px 0 15px;font:800 clamp(3.2rem,7vw,6.4rem)/.9 'Playfair Display',serif;letter-spacing:-3px;color:#092536;}
.hero h1 .orange {color:#e85d1b;}
.hero h1 .green {color:#08724e;}
.hero-desc {font:500 1.25rem/1.55 'Caveat',cursive;color:#29333a;max-width:590px;}
.hero-art {position:absolute;right:20px;bottom:10px;width:47%;height:82%;border-radius:48% 0 0 48%;background:linear-gradient(150deg,rgba(9,46,67,.02),rgba(9,46,67,.15));}
.hero-art .gate {position:absolute;right:19%;bottom:4%;width:185px;height:300px;border:25px solid #c77b43;border-bottom-width:12px;border-radius:12px 12px 0 0;box-shadow:inset 0 0 0 5px #e2a36d;}
.hero-art .gate:before {content:"INDIA";position:absolute;top:-18px;left:22px;width:90px;text-align:center;background:#c77b43;color:#754522;font:bold 13px Inter;padding:5px;}
.hero-art .gate:after {content:"";position:absolute;left:30px;right:30px;top:55px;bottom:0;border:18px solid #d28b52;border-bottom:0;border-radius:90px 90px 0 0;}
.temple {position:absolute;right:2%;bottom:2%;font-size:9rem;filter:sepia(.25);opacity:.9;}

.categories {display:grid;grid-template-columns:repeat(6,1fr);gap:14px;padding:8px 42px 34px;background:#fff7eb;}
.cat-link{display:block;color:inherit;text-decoration:none}.cat {text-align:center;cursor:pointer;transition:.18s}.cat-link:hover .cat{transform:translateY(-4px);box-shadow:0 14px 28px rgba(20,45,55,.16)}
.cat-old {text-align:center;border:2px solid rgba(255,255,255,.95);border-radius:24px;padding:17px 10px 15px;box-shadow:0 9px 22px rgba(20,45,55,.09);font-family:Inter,sans-serif;}
.cat:nth-child(1){background:#fff0e5}.cat:nth-child(2){background:#fdecef}.cat:nth-child(3){background:#fff5d9}.cat:nth-child(4){background:#e9f4ed}.cat:nth-child(5){background:#fde7e5}.cat:nth-child(6){background:#e5f0f4}
.cat-icon {font-size:2.7rem;line-height:1.05}.cat-name{font-weight:800;font-size:1.1rem;margin-top:8px}.cat-desc{font-size:.78rem;line-height:1.2;color:#555;margin-top:5px}

.search-hero {margin:0 42px 28px;padding:4px;border-radius:38px;background:#fff;box-shadow:0 9px 28px rgba(24,35,40,.13);display:flex;align-items:center;}
.search-hero-link{display:flex;align-items:center;flex:1;color:inherit;text-decoration:none}.search-hero .search-icon {font-size:1.8rem;padding:0 14px;color:#777}.search-hero .search-text {flex:1;color:#6a6a6a;font:500 1rem Inter,sans-serif;padding:14px 0}.search-hero .search-arrow {width:64px;height:52px;border-radius:30px;background:#e96922;color:white;display:grid;place-items:center;font-size:1.8rem;margin-right:4px}

.explore-band {background:#062a43;color:white;padding:38px 42px 45px;position:relative;overflow:hidden;}
.explore-band:after {content:"";position:absolute;right:-80px;top:-100px;width:300px;height:300px;border:2px solid rgba(255,255,255,.08);border-radius:50%;}
.explore-title {font:800 1.75rem Inter,sans-serif}.map-pill-link{color:inherit;text-decoration:none}.feature-link{display:block;color:inherit;text-decoration:none}.feature{transition:.18s}.feature-link:hover .feature{background:rgba(255,255,255,.13);transform:translateY(-3px)}.explore-sub{margin-top:5px;color:#d9e7ef}.map-pill{float:right;margin-top:-44px;border:2px solid #4eb184;border-radius:28px;padding:11px 22px;color:#a9e3c7;font-weight:800}
.feature-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:18px;margin-top:28px}.feature{background:rgba(255,255,255,.07);border:1px solid rgba(255,255,255,.18);border-radius:22px;padding:23px;display:flex;gap:16px;align-items:center}.feature-icon{font-size:2.1rem}.feature h3{margin:0;font-size:1rem}.feature p{margin:6px 0 0;color:#c8d7df;font-size:.84rem}
.home-footer {padding:34px 20px 45px;text-align:center;background:#fff1dd;color:#173246}.unity{letter-spacing:4px;font:600 .78rem Inter,sans-serif}.quote{margin-top:22px;font:600 1.1rem 'Playfair Display',serif}.tricolor{margin:17px auto 0;width:190px;height:4px;background:linear-gradient(90deg,#f07b28 0 33%,#f8f1df 33% 66%,#2d8a61 66%);border-radius:4px}

@media(max-width:900px){.nav-links{display:none}.motto{font-size:1rem}.home-nav{padding:16px 18px}.hero{padding:38px 24px 25px;min-height:560px}.hero-art{opacity:.32;width:80%;right:-70px}.categories{grid-template-columns:repeat(3,1fr);padding:8px 18px 24px}.search-hero{margin:0 18px 24px}.explore-band{padding:30px 20px}.feature-grid{grid-template-columns:1fr}.map-pill{float:none;display:inline-block;margin:20px 0 0}.unity{letter-spacing:2px}}
@media(max-width:560px){.brand-name{font-size:1.25rem}.brand-sub{font-size:.55rem;letter-spacing:2.5px}.hero h1{font-size:3.45rem;letter-spacing:-1.5px}.hero-desc{font-size:1.12rem}.categories{grid-template-columns:repeat(2,1fr);gap:10px}.cat{padding:14px 6px}.cat-icon{font-size:2.2rem}.hero-art{display:none}.search-hero .search-text{font-size:.88rem}.nav-links{display:none}}
</style>
""", unsafe_allow_html=True)

# Real HTML/CSS frontend: every section is an actual DOM element, not one large image.
st.markdown("""
<div class="home-front">
  <div class="home-nav">
    <div class="brand"><div class="brand-flag">🇮🇳</div><div><div class="brand-name"><span>SMART</span> INDIA</div><div class="brand-sub">CULTURAL DISCOVERY</div></div></div>
    <div class="nav-links"><a href="?home_action=map">⌂ Home</a><a href="?home_action=map">🗺 Explore Map</a><a href="?home_action=about">ⓘ About</a><a href="?home_action=chat">💬 Chat</a></div>
    <div class="motto">Our Culture<br>Our Identity<br>Our Pride<i></i></div>
  </div>

  <section class="hero">
    <div class="hero-copy">
      <div class="kicker">🇮🇳 SMART INDIA • CULTURAL DISCOVERY</div>
      <h1>Discover<br><span class="orange">India's</span><br><span class="green">Cultural DNA</span></h1>
      <div class="hero-desc">Explore the living traditions of India through state-wise stories, authentic cultural visuals, relevant videos and interactive maps — all from one place.</div>
    </div>
    <div class="hero-art"><div class="gate"></div><div class="temple">🛕</div></div>
  </section>

  <div class="categories">
    <a class="cat-link" href="?home_prompt=Food+of+Indian+states"><div class="cat"><div class="cat-icon">🍛</div><div class="cat-name">Food</div><div class="cat-desc">Flavours<br>of India</div></div></a>
    <a class="cat-link" href="?home_prompt=Traditional+dress+of+Indian+states"><div class="cat"><div class="cat-icon">👗</div><div class="cat-name">Dress</div><div class="cat-desc">Traditions<br>in Threads</div></div></a>
    <a class="cat-link" href="?home_prompt=Folk+dances+of+Indian+states"><div class="cat"><div class="cat-icon">💃</div><div class="cat-name">Dance</div><div class="cat-desc">Rhythms<br>of Heritage</div></div></a>
    <a class="cat-link" href="?home_prompt=Folk+music+of+Indian+states"><div class="cat"><div class="cat-icon">🎵</div><div class="cat-name">Music</div><div class="cat-desc">Melodies<br>of Culture</div></div></a>
    <a class="cat-link" href="?home_prompt=Festivals+of+Indian+states"><div class="cat"><div class="cat-icon">🪔</div><div class="cat-name">Festivals</div><div class="cat-desc">Celebrations<br>of Unity</div></div></a>
    <a class="cat-link" href="?home_prompt=Traditional+crafts+of+Indian+states"><div class="cat"><div class="cat-icon">🏺</div><div class="cat-name">Crafts</div><div class="cat-desc">Art in<br>Every Hand</div></div></a>
  </div>

  <div class="search-hero"><a class="search-hero-link" href="?home_prompt=Tell%20me%20about%20Indian%20culture"><div class="search-icon">⌕</div><div class="search-text">Ask about a state, city, capital, monument, festival...</div><div class="search-arrow">→</div></a></div>

  <section class="explore-band">
    <div class="explore-title">🗺️ Explore India on Map</div>
    <div class="explore-sub">Click on any state to discover its unique culture</div>
    <a class="map-pill-link" href="?home_action=map"><div class="map-pill">View Map →</div></a>
    <div class="feature-grid">
      <a class="feature-link" href="?home_prompt=Show+me+state-wise+cultural+information"><div class="feature"><div class="feature-icon">📖</div><div><h3>State-wise Information</h3><p>Detailed cultural insights</p></div></div></a>
      <a class="feature-link" href="?home_prompt=Show+me+authentic+cultural+images"><div class="feature"><div class="feature-icon">🖼️</div><div><h3>Authentic Images</h3><p>With source &amp; credit</p></div></div></a>
      <a class="feature-link" href="?home_prompt=Show+me+relevant+cultural+videos"><div class="feature"><div class="feature-icon">▶️</div><div><h3>Relevant Videos</h3><p>Watch real cultural content</p></div></div></a>
    </div>
  </section>

  <div class="home-footer"><div class="unity">ONE NATION • MANY CULTURES • A SHARED TOMORROW</div><div class="tricolor"></div><div class="quote">“India is not just a land, it’s a living culture.”</div></div>
</div>
""", unsafe_allow_html=True)

# Sidebar: State Selection & Quick Filters
with st.sidebar:
    st.header("📍 Select a Region")
    state_options = ["Select a state/region"] + list_states()
    selected = st.selectbox("Choose State / UT", state_options, key="sidebar_state_select")
    voice_language = st.selectbox(
        "🎤 Voice Language",
        ["English (India)", "Hindi (India)"],
        index=0,
        key="voice_language"
    )
    voice_language_code = "en-IN" if voice_language.startswith("English") else "hi-IN"

    active_state = None if selected == "Select a state/region" else selected

    if active_state:
        pretty = get_display_name(active_state)
        st.success(f"**Selected:** {pretty}")

        state_meta = STATE_COORDINATES.get(active_state)
        if state_meta:
            st.markdown(f"🏛️ **Capital / Center:** {state_meta.get('capital', 'N/A')}")
            st.markdown(f"📍 **Coordinates:** `{state_meta['lat']:.2f}° N, {state_meta['lon']:.2f}° E`")

        st.markdown("### ⚡ Quick Explore")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🍲 Cuisine", use_container_width=True):
                st.session_state["queued_prompt"] = f"Traditional food of {pretty}"
            if st.button("💃 Dances", use_container_width=True):
                st.session_state["queued_prompt"] = f"Folk dances of {pretty}"
            if st.button("👗 Clothing", use_container_width=True):
                st.session_state["queued_prompt"] = f"Traditional attire of {pretty}"
        with col2:
            if st.button("🏛️ Heritage", use_container_width=True):
                st.session_state["queued_prompt"] = f"Heritage and monuments of {pretty}"
            if st.button("🎉 Festivals", use_container_width=True):
                st.session_state["queued_prompt"] = f"Festivals and celebrations of {pretty}"
            if st.button("🗺️ Show Map", use_container_width=True):
                st.session_state["queued_prompt"] = f"Show {pretty} on map"

    st.divider()
    st.markdown("### 💡 Popular Prompts")
    st.markdown(
        "- *What is the traditional food of Jammu?*\n"
        "- *Festivals of Kashmir*\n"
        "- *Tell me about Rajasthan*\n"
        "- *Traditional dress of Punjab*\n"
        "- *Folk dance of Gujarat*\n"
        "- *Where is Kerala on map?*"
    )
    st.caption("Culture DNA Bot • Powered by Folium & Wikimedia")

# =====================================================================
# 10. STATE SPOTLIGHT CARD (Requirement 2 & 8: When user selects a state)
# Shows: 1. State Name, 2. Culture Info, 3. One relevant Image, 4. Folium Map
# =====================================================================
if active_state:
    row = row_for_state(active_state)
    if row is not None:
        display_name = get_display_name(active_state)
        visual_data = get_visual_for_reply(active_state, None, row)

        with st.container():
            st.markdown('<div class="state-spotlight-card">', unsafe_allow_html=True)
            st.markdown(f'<div class="state-spotlight-title">🏛️ {display_name} — Culture & Heritage</div>', unsafe_allow_html=True)
            
            # Culture Information summary
            points = []
            for col in [
                "Food & Cuisine",
                "Folk Dances",
                "Festivals & Celebrations",
                "Heritage & Monuments",
                "Traditional Clothing",
                "Languages & Dialects"
            ]:
                val = str(row.get(col, "")).strip()
                if val:
                    points.append(f"- **{FIELD_LABELS.get(col, col)}:** {val}")
            
            if points:
                st.markdown("\n".join(points))

            st.write("")
            col_img, col_map = st.columns([1, 1])
            with col_img:
                render_visual_gallery(visual_data)
            
            with col_map:
                coords = visual_data.get("coords")
                if coords:
                    st.markdown(f"**🗺️ Interactive Map (Folium):** *{visual_data.get('location_name', display_name)}*")
                    # Renders interactive Folium map automatically centered on this state
                    render_interactive_map(
                        coords["lat"],
                        coords["lon"],
                        visual_data.get("location_name", display_name),
                        map_key=f"spotlight_{active_state}_{coords['lat']}_{coords['lon']}"
                    )
            st.markdown('</div>', unsafe_allow_html=True)

# =====================================================================
# 11. CHATBOT INTERACTION (Requirement 2: When user asks about a state)
# =====================================================================
st.markdown("### 💬 Ask Culture DNA")

# Initialize Chat History
if "messages" not in st.session_state:
    st.session_state.messages = [
        {
            "role": "assistant",
            "content": (
                "Namaste! 👋 I am **Culture DNA**.\n\n"
                "Ask me about any Indian state's culture, food, folk dances, or monuments. "
                "I will answer with cultural details, **an authentic photo**, and **an interactive map**!"
            ),
            "visual": None
        }
    ]

# Render existing chat history
for idx, msg in enumerate(st.session_state.messages):
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

        visual = msg.get("visual")
        if visual:
            col_img, col_map = st.columns([1.1, 1.1])
            with col_img:
                render_visual_gallery(visual)
            with col_map:
                if visual.get("coords"):
                    coords = visual["coords"]
                    loc_name = visual.get("location_name", "Location")
                    st.markdown(f'<div class="map-header">📍 <b>Location Pin:</b> {loc_name}</div>', unsafe_allow_html=True)
                    render_interactive_map(
                        coords["lat"],
                        coords["lon"],
                        loc_name,
                        map_key=f"chat_map_{idx}_{coords['lat']}_{coords['lon']}"
                    )

# Handle Chat Input, Voice Input & Queued Prompts
queued = st.session_state.pop("queued_prompt", None)

voice_prompt = None
voice_error = False
with st.container():
    st.markdown('<div class="voice-box">🎤 <b>Voice Search</b><div class="voice-help">Tap Record, speak your question, then tap Stop. You can ask in English or Hindi.</div></div>', unsafe_allow_html=True)
    voice_audio = mic_recorder(
        start_prompt="🎤 Record",
        stop_prompt="⏹️ Stop",
        just_once=True,
        format="wav",
        key="culture_dna_voice"
    )
    if voice_audio:
        with st.spinner("🎤 Converting your voice to text..."):
            transcript = transcribe_voice(voice_audio, voice_language_code)
        if transcript == "__VOICE_SERVICE_ERROR__":
            voice_error = True
            st.error("Voice recognition service is unavailable right now. Please use text input.")
        elif transcript:
            voice_prompt = transcript
            st.success(f"🎤 You said: **{transcript}**")
        else:
            st.warning("I could not understand the recording. Please speak clearly and try again.")

user_prompt = st.chat_input("Ask about a state, city, capital, monument, dance, or food…")

# Typed text takes priority; otherwise use voice; otherwise use queued home/sidebar action.
active_prompt = user_prompt or voice_prompt or queued

if active_prompt:
    selected_state = None if selected == "Select a state/region" else selected

    # 1. Append User message
    st.session_state.messages.append({
        "role": "user",
        "content": active_prompt,
        "visual": None
    })

    # 2. Generate Assistant Reply
    reply_obj = chatbot_reply(active_prompt, selected_state)

    # 3. Append Assistant Message with Visual & Map Data
    st.session_state.messages.append({
        "role": "assistant",
        "content": reply_obj["text"],
        "visual": reply_obj["visual"]
    })

    # 4. Clean rerun to sync state and display images & maps
    st.rerun()
