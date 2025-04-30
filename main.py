import os
import random
import requests
import io
import json
from pptx import Presentation
from pptx.util import Inches
from PIL import Image
import google.generativeai as genai

# --- CONFIG ---
GOOGLE_API_KEY = 'AIzaSyAABoqmArgjddo1ROP9q4dLYf_Bo7WVLsQ'        
GOOGLE_CX = '35d98aa59458c428a'          
GEMINI_API_KEY = 'AIzaSyAVwOIb2B3hE_hRxlpt5EWcNANAsg3eJ8U'

# --- SETUP ---
genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel("models/gemini-1.5-pro-latest")

# --- IMAGE CONVERSION ---
def convert_image_to_supported_format(image_bytes):
    img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    output = io.BytesIO()
    img.save(output, format="PNG")
    output.seek(0)
    return output

# --- STRUCTURED SUBTOPIC GENERATION ---
def generate_subtopics(topic, count=5):
    prompt = f"""
Generate a list of {count} presentation slide objects in JSON.
Each object should include a 'title' and 'content' field.
The topic is: {topic}

Example:
[
  {{"title": "Introduction to AI", "content": "Artificial Intelligence is the simulation of human intelligence..."}},
  ...
]
Only return valid JSON.
"""
    response = model.generate_content(prompt)
    try:
        slides = json.loads(response.text)
        return [(s["title"], s["content"]) for s in slides if "title" in s and "content" in s]
    except json.JSONDecodeError:
        print("Error: Gemini did not return valid JSON.")
        return []

# --- GOOGLE IMAGE FETCH ---
def get_image_url(query):
    url = f"https://www.googleapis.com/customsearch/v1?q={query}&cx={GOOGLE_CX}&key={GOOGLE_API_KEY}&searchType=image&num=1"
    try:
        res = requests.get(url).json()
        return res["items"][0]["link"]
    except:
        return None

# --- ALTERNATING IMAGE POSITIONS ---
def alternating_image_position(index):
    positions = [
        (Inches(5.5), Inches(0.5), Inches(3)),  # Top-right
        (Inches(0.5), Inches(4.5), Inches(3.5)),  # Bottom-left
    ]
    return positions[index % len(positions)]

# --- SLIDE CREATION ---
def create_content_slide(prs, title, content, image_stream=None, index=0):
    layout = prs.slide_layouts[1]  # Title + Content
    slide = prs.slides.add_slide(layout)
    slide.shapes.title.text = title
    slide.placeholders[1].text = content

    if image_stream:
        left, top, width = alternating_image_position(index)
        slide.shapes.add_picture(image_stream, left, top, width=width)

# --- TEMPLATE LOADER ---
def load_random_template():
    templates = [f for f in os.listdir("templates") if f.endswith(".pptx")]
    if not templates:
        raise FileNotFoundError("No templates found in the 'templates' folder.")
    return Presentation(os.path.join("templates", random.choice(templates)))

# --- MAIN APP ---
def main():
    topic = input("Enter your presentation topic: ")
    slide_count = int(input("How many slides? "))

    subtopics = generate_subtopics(topic, count=slide_count)
    if not subtopics:
        print("Failed to generate slides. Exiting.")
        return

    prs = load_random_template()

    for i, (title, content) in enumerate(subtopics):
        image_url = get_image_url(title)
        image_stream = None
        if image_url:
            try:
                image_data = requests.get(image_url).content
                image_stream = convert_image_to_supported_format(image_data)
            except:
                print(f"Image load failed for: {title}")
        create_content_slide(prs, title, content, image_stream, i)

    filename = f"{topic.strip().replace(' ', '_')}_presentation.pptx"
    prs.save(filename)
    print(f"✅ Presentation saved as: {filename}")

if __name__ == "__main__":
    main()
