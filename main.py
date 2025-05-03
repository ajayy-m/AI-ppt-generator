import os
import json
import random
import requests
import io
import re
from pptx import Presentation
from pptx.util import Inches
from google.generativeai import configure, GenerativeModel
from PIL import Image

# === CONFIGURATION ===
GOOGLE_API_KEY = "AIzaSyAABoqmArgjddo1ROP9q4dLYf_Bo7WVLsQ"  # Replace this with your actual Gemini API key
SEARCH_ENGINE_ID = "35d98aa59458c428a"  # Replace with your Google Custom Search Engine ID
GEMINI_API = "AIzaSyAVwOIb2B3hE_hRxlpt5EWcNANAsg3eJ8U"

configure(api_key=GEMINI_API)
model = GenerativeModel(model_name="models/gemini-1.5-pro")

TEMPLATE_DIR = "templates"

# === UTILS ===
def extract_valid_json(text):
    match = re.search(r'\[.*\]', text, re.DOTALL)
    if match:
        try:
            return json.loads(match.group())
        except json.JSONDecodeError:
            return []
    return []

def generate_subtopics(topic, count=5, retries=2):
    assert count >= 3, "Slide count must be at least 3 to include intro and conclusion."
    middle = count - 2
    prompt = f'''
You are an expert presentation assistant.

Generate a JSON list of {count} slides on the topic: "{topic}", structured as follows:
1. "Introduction" — a slide introducing the topic
2. {middle} subtopic slides (AI-generated subtopics relevant to the topic)
3. "Conclusion" — a summary or concluding slide

Format:
[
  {{
    "title": "Slide Title",
    "content": "Slide content..."
  }},
  ...
]

Output ONLY the JSON list, no extra text or commentary.
Ensure the JSON is valid.
'''
    for attempt in range(retries):
        response = model.generate_content(prompt)
        slides = extract_valid_json(response.text)
        if slides:
            return [(s["title"], s["content"]) for s in slides if "title" in s and "content" in s]
        print(f"⚠️ Attempt {attempt + 1}: Failed to get valid JSON from Gemini.")
    print("❌ Gemini failed to return valid slide data after retries.")
    return []

def fetch_image(main_topic, subtopic):
    query = f"{main_topic} {subtopic}"
    search_url = f"https://www.googleapis.com/customsearch/v1?q={query}&searchType=image&key={GOOGLE_API_KEY}&cx={SEARCH_ENGINE_ID}&num=1"
    response = requests.get(search_url).json()
    items = response.get("items")
    if items:
        img_url = items[0]["link"]
        try:
            img_data = requests.get(img_url).content
            if img_url.lower().endswith(".webp"):
                img = Image.open(io.BytesIO(img_data)).convert("RGB")
                buf = io.BytesIO()
                img.save(buf, format="PNG")
                return buf
            return io.BytesIO(img_data)
        except Exception as e:
            print(f"⚠️ Error fetching image: {e}")
    return None

def create_content_slide(prs, title, content, image_stream, index):
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)

    left_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(4.5), Inches(6))
    right_box = slide.shapes.add_textbox(Inches(5), Inches(0.5), Inches(4.5), Inches(6))

    if index % 2 == 0:
        if image_stream:
            slide.shapes.add_picture(image_stream, Inches(0.5), Inches(1), width=Inches(4.5))
        tf = right_box.text_frame
    else:
        tf = left_box.text_frame
        if image_stream:
            slide.shapes.add_picture(image_stream, Inches(5), Inches(1), width=Inches(4.5))

    tf.text = title
    p = tf.add_paragraph()
    p.text = content
    p.level = 1

def add_title_slide(prs, topic):
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    slide.shapes.title.text = topic
    slide.placeholders[1].text = "Your Name Here"

def choose_template():
    templates = [f for f in os.listdir(TEMPLATE_DIR) if f.endswith(".pptx")]
    if not templates:
        raise FileNotFoundError("No PPTX templates found in 'templates/' folder.")
    print("Available templates:")
    for idx, t in enumerate(templates, 1):
        print(f"{idx}. {t}")
    choice = int(input("Choose a template by number: ")) - 1
    return os.path.join(TEMPLATE_DIR, templates[choice])

# === MAIN ===
def main():
    topic = input("Enter your presentation topic: ")
    slide_count = int(input("How many slides (min 3)? "))

    subtopics = generate_subtopics(topic, count=slide_count)
    if not subtopics:
        print("❌ Failed to generate slides.")
        return

    template_path = choose_template()
    prs = Presentation(template_path)

    add_title_slide(prs, topic)

    for idx, (title, content) in enumerate(subtopics):
        img = fetch_image(topic, title)
        create_content_slide(prs, title, content, img, idx)

    output_path = f"{topic.replace(' ', '_')}.pptx"
    prs.save(output_path)
    print(f"✅ Presentation saved to {output_path}")

if __name__ == "__main__":
    main()
