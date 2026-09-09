import streamlit as st
import pandas as pd
import re
from pathlib import Path

st.set_page_config(
    page_title="Culture DNA",
    page_icon="🧬",
    layout="wide"
)

DATA_FILE = Path(__file__).parent / "Culture_DNA.xlsx"

@st.cache_data
def load_data():
    df = pd.read_excel(DATA_FILE)
    df = df.fillna("")
    return df

df = load_data()

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

def clean_state_name(name):
    name = str(name).strip()
    if name in df["State/UT"].tolist():
        return name
    low = name.lower()
    if low in ALIASES:
        return ALIASES[low]
    for state in df["State/UT"]:
        if low == str(state).lower():
            return state
    return None

def find_state(text):
    text_low = text.lower()
    # Prefer Kashmir before Jammu because "Jammu and Kashmir" should map to the
    # Jammu-region entry unless the user explicitly asks for Kashmir.
    if "kashmir valley" in text_low or re.search(r"\bkashmir\b", text_low):
        if "jammu" not in text_low:
            return "1.2 Kashmir"
    if "jammu" in text_low or "j&k" in text_low:
        return "1.1 Jammu"

    for state in df["State/UT"]:
        state_low = str(state).lower()
        if state_low in text_low:
            return state

    for alias, state in ALIASES.items():
        if alias in text_low:
            return state

    return None

def find_field(text):
    text_low = text.lower()
    # Long phrases first.
    for key, field in sorted(QUESTION_MAP.items(), key=lambda x: len(x[0]), reverse=True):
        if key in text_low:
            return field
    return None

def row_for_state(state):
    rows = df[df["State/UT"].astype(str) == state]
    return rows.iloc[0] if not rows.empty else None

def list_states():
    return [str(x) for x in df["State/UT"].tolist()]

def chatbot_reply(message, selected_state=None):
    text = message.strip()
    low = text.lower()

    if not text:
        return "Please ask me something about a state's culture. 😊"

    if any(x in low for x in ["hello", "hi", "hey", "namaste"]):
        return (
            "Namaste! 👋 I am **Culture DNA**, your state-wise culture chatbot. "
            "Ask me about food, festivals, languages, dance, clothing, heritage, "
            "crafts, traditions and more."
        )

    if "which states" in low or "list states" in low or "all states" in low:
        return "I currently have these entries:\n\n" + "\n".join(
            f"- {s}" for s in list_states()
        )

    state = find_state(text) or selected_state
    field = find_field(text)

    if state:
        row = row_for_state(state)
        if row is None:
            return "I couldn't find that state in the dataset."

        display_name = state.replace("1.1 ", "Jammu — ").replace("1.2 ", "Kashmir — ")

        if field:
            value = str(row[field]).strip()
            if not value:
                return f"I don't have verified data for **{FIELD_LABELS[field]}** for {display_name} yet."
            return f"### {display_name}\n\n**{FIELD_LABELS[field]}:**\n{value}"

        # General state profile.
        parts = []
        for col in [
            "Languages & Dialects",
            "Folk Dances",
            "Folk Music",
            "Traditional Clothing",
            "Food & Cuisine",
            "Festivals & Celebrations",
            "Heritage & Monuments",
            "Arts & Crafts",
        ]:
            value = str(row[col]).strip()
            if value:
                parts.append(f"**{FIELD_LABELS[col]}:** {value}")

        return f"## 🧬 Culture DNA — {display_name}\n\n" + "\n\n".join(parts)

    if field:
        return (
            f"I can answer about **{FIELD_LABELS[field]}**. "
            "Please include a state/UT name, for example: "
            "**food of Jammu** or **festivals of Kashmir**."
        )

    return (
        "I can help with:\n\n"
        "- 🍲 Food & Cuisine\n"
        "- 🎉 Festivals & Celebrations\n"
        "- 💃 Folk Dances\n"
        "- 🎵 Folk Music\n"
        "- 🗣️ Languages & Dialects\n"
        "- 👗 Traditional Clothing\n"
        "- 🏛️ Heritage & Monuments\n"
        "- 🎨 Arts & Crafts\n"
        "- 🧿 Traditions, communities, architecture, textiles and more\n\n"
        "Try: **What is the traditional food of Jammu?**"
    )

# ---------- UI ----------
st.title("🧬 Culture DNA")
st.caption("Explore India's culture state by state — powered by the Culture DNA dataset.")

with st.sidebar:
    st.header("Choose a region")
    state_options = ["Select a state/region"] + list_states()
    selected = st.selectbox("State / UT", state_options)

    if selected != "Select a state/region":
        pretty = selected.replace("1.1 ", "Jammu — ").replace("1.2 ", "Kashmir — ")
        st.success(f"Selected: {pretty}")

    st.divider()
    st.markdown("### Try asking")
    st.markdown(
        "- Food of Jammu\n"
        "- Festivals of Kashmir\n"
        "- Traditional dress of Punjab\n"
        "- Folk dance of Rajasthan\n"
        "- Heritage of Himachal Pradesh"
    )

if "messages" not in st.session_state:
    st.session_state.messages = [
        {
            "role": "assistant",
            "content": (
                "Namaste! 👋 I'm **Culture DNA**. "
                "Ask me anything about the culture of a state or region."
            ),
        }
    ]

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

prompt = st.chat_input("Ask about a state's culture…")

if prompt:
    selected_state = None if selected == "Select a state/region" else selected

    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    reply = chatbot_reply(prompt, selected_state)

    st.session_state.messages.append({"role": "assistant", "content": reply})
    with st.chat_message("assistant"):
        st.markdown(reply)
