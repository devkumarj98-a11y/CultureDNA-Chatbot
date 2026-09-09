import streamlit as st
import pandas as pd
import requests
import re
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
    initial_sidebar_state="expanded"
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
    .stChatMessage {
        border-radius: 14px;
        padding: 12px;
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
    text_low = text.lower()
    for key, field in sorted(QUESTION_MAP.items(), key=lambda x: len(x[0]), reverse=True):
        if re.search(rf"\b{re.escape(key)}\b", text_low):
            return field
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

def extract_visual_topic(text: str, field: str = None):
    """Extract the user's requested subject so image search follows the query."""
    low = text.lower().strip()
    glue = r"\b(please|show|me|tell|about|what|which|is|are|the|of|in|from|for|on|can|you|give|information|info|details|picture|photo|image|images|related|relevant|famous|popular|traditional|culture|cultural)\b"
    cleaned = re.sub(glue, " ", low, flags=re.IGNORECASE)

    known_locations = list(LOCATION_ALIASES.keys()) + list(STATE_COORDINATES.keys()) + [
        get_display_name(x) for x in STATE_COORDINATES.keys()
    ]
    for name in sorted(set(known_locations), key=len, reverse=True):
        if name:
            cleaned = re.sub(rf"\b{re.escape(name.lower())}\b", " ", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"[^a-z0-9&' -]", " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()

    if field:
        label = FIELD_LABELS.get(field, field)
        return f"{cleaned} {label}".strip() if cleaned else label
    return cleaned or "culture"


def get_visual_for_reply(state: str, field: str = None, row=None, location: str = None, location_coords=None, user_query: str = ""):
    """Return an image matching the user's requested subject, plus map/video."""
    state_meta = STATE_COORDINATES.get(state, {})
    clean_name = get_clean_region_name(state)
    display_title = get_display_name(state)
    location_label = location.strip() if location else display_title

    topic_text = extract_visual_topic(user_query, field)
    topic_visual_result = None

    # Highest priority: what the user actually asked for + where.
    topic_queries = [
        f"{topic_text} {location_label} India",
        f"{location_label} {topic_text} India",
    ]

    # Dataset-specific examples are fallback searches, not the first choice.
    if field and row is not None:
        cell_val = str(row.get(field, "")).strip()
        first_item = re.split(r"[,;:\n(]", cell_val)[0].strip()
        first_item = re.sub(
            r"^(include|including|such as|major|famous|popular)\s+",
            "", first_item, flags=re.IGNORECASE
        ).strip()
        if len(first_item) > 2:
            topic_queries.extend([
                f"{first_item} {location_label} India",
                f"{first_item} India {FIELD_LABELS.get(field, field)}",
            ])

    for q in topic_queries:
        result = fetch_cultural_visual(q)
        if result and result.get("image_url"):
            topic_visual_result = result
            break

    # Generic location/state image is used only when no specific subject was requested.
    custom_image_url = None
    if not field and row is not None and "image_url" in row:
        val = str(row["image_url"]).strip()
        if val.startswith(("http://", "https://")):
            custom_image_url = val

    if not field and not custom_image_url and row is not None and "Multimedia / Images" in row:
        val = str(row["Multimedia / Images"]).strip()
        match = re.search(r"https?://[^\s,;]+", val)
        if match:
            custom_image_url = match.group(0)

    if not field and not custom_image_url and state in STATE_CUSTOM_IMAGE_URLS:
        val = STATE_CUSTOM_IMAGE_URLS[state].strip()
        if val.startswith(("http://", "https://")):
            custom_image_url = val

    visual_result = topic_visual_result
    if not custom_image_url and not visual_result and location:
        visual_result = fetch_cultural_visual(f"{location} {clean_name} India")
        if not visual_result:
            visual_result = fetch_cultural_visual(location)
    if not custom_image_url and not visual_result and state in STATE_COORDINATES:
        visual_result = fetch_cultural_visual(state_meta.get("query", f"{clean_name} landmark"))
    if not custom_image_url and not visual_result:
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

    video_query = f"{location_label} {topic_text}" if (field or topic_text != "culture") else f"{location_label} {clean_name} India culture"
    return {
        "image_url": custom_image_url or (visual_result.get("image_url") if visual_result else None),
        "image_caption": visual_result.get("title") if visual_result else location_label,
        "coords": final_coords,
        "location_name": location_name,
        "video_url": youtube_search_url(video_query),
        "video_query": video_query,
    }

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

        parts = []
        for col in [
            "Languages & Dialects", "Folk Dances", "Folk Music",
            "Traditional Clothing", "Food & Cuisine",
            "Festivals & Celebrations", "Heritage & Monuments", "Arts & Crafts"
        ]:
            value = str(row.get(col, "")).strip()
            if value:
                parts.append(f"**{FIELD_LABELS[col]}:** {value}")

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
# 9. UI LAYOUT & DISPLAY
# =====================================================================
st.markdown('<div class="main-title">🧬 Culture DNA</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Explore India\'s rich culture and heritage by state, capital, city or region with authentic photos, videos & interactive maps.</div>', unsafe_allow_html=True)

# Sidebar: State Selection & Quick Filters
with st.sidebar:
    st.header("📍 Select a Region")
    state_options = ["Select a state/region"] + list_states()
    selected = st.selectbox("Choose State / UT", state_options, key="sidebar_state_select")

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
                st.markdown("**📸 Cultural Image:**")
                if visual_data.get("image_url"):
                    st.image(
                        visual_data["image_url"],
                        caption=f"Visual: {visual_data.get('image_caption', display_name)}",
                        use_container_width=True
                    )
                else:
                    st.info("No image found. You can add one in STATE_CUSTOM_IMAGE_URLS.")

                if visual_data.get("video_url"):
                    st.markdown("**🎥 Cultural Video:**")
                    st.link_button(
                        "▶️ Find a relevant cultural video",
                        visual_data["video_url"],
                        use_container_width=True
                    )
                    st.caption(f"Search: {visual_data.get('video_query', '')}")
            
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
                if visual.get("image_url"):
                    st.image(
                        visual["image_url"],
                        caption=f"📸 {visual.get('image_caption', 'Cultural Heritage')}",
                        use_container_width=True
                    )
                if visual.get("video_url"):
                    st.markdown("**🎥 Cultural Video**")
                    st.link_button(
                        "▶️ Find relevant video",
                        visual["video_url"],
                        use_container_width=True
                    )
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

# Handle Chat Input & Queued Prompts
queued = st.session_state.pop("queued_prompt", None)
user_prompt = st.chat_input("Ask about a state, city, capital, monument, dance, or food…")

active_prompt = queued or user_prompt

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
