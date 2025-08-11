import streamlit as st
import os
import json
import re
from typing import List, Dict, Any, Optional
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from datetime import datetime
import io
import base64
import requests
from PIL import Image, ImageDraw, ImageFont
import time

# Try to import Google Generative AI, fallback if not available
try:
    import google.generativeai as genai
    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False
    st.warning("⚠️ Google Generative AI not available. Using fallback content generation.")

# Set page config
st.set_page_config(
    page_title="AI PowerPoint Generator Pro",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
.main {
    padding-top: 1rem;
}
.stButton>button {
    width: 100%;
    border-radius: 10px;
    height: 3em;
}
.success-message {
    padding: 1rem;
    border-radius: 10px;
    background-color: #d4edda;
    border: 1px solid #c3e6cb;
    color: #155724;
}
.info-box {
    padding: 1rem;
    border-radius: 10px;
    background-color: #e3f2fd;
    border: 1px solid #bbdefb;
    color: #0d47a1;
    margin: 1rem 0;
}
.custom-content-box {
    background-color: #f8f9fa;
    border-radius: 10px;
    padding: 1rem;
    border-left: 4px solid #007bff;
}
.error-message {
    padding: 1rem;
    border-radius: 10px;
    background-color: #f8d7da;
    border: 1px solid #f5c6cb;
    color: #721c24;
}
</style>
""", unsafe_allow_html=True)

class GeminiContentGenerator:
    """Enhanced AI content generator using Gemini AI"""
    
    def __init__(self, api_key: str = None):
        self.api_key = api_key or GEMINI_API_KEY
        self.model = None
        
        if GEMINI_AVAILABLE and self.api_key:
            try:
                genai.configure(api_key=self.api_key)
                self.model = genai.GenerativeModel('gemini-pro')
                st.success("🤖 Gemini AI initialized successfully!")
            except Exception as e:
                st.error(f"❌ Failed to initialize Gemini AI: {str(e)}")
                self.model = None
    
    def generate_content_structure(self, topic: str, num_slides: int = 5, 
                                 focus_areas: List[str] = None, 
                                 custom_content: str = None,
                                 target_audience: str = None,
                                 presentation_style: str = "professional") -> Dict[str, Any]:
        """Generate AI-powered presentation structure using Gemini"""
        
        if self.model and GEMINI_AVAILABLE:
            try:
                return self._generate_with_gemini(topic, num_slides, focus_areas, custom_content, target_audience, presentation_style)
            except Exception as e:
                st.warning(f"⚠️ Gemini AI error: {str(e)}. Using fallback generation.")
                return self._generate_fallback_content(topic, num_slides, focus_areas, custom_content)
        else:
            return self._generate_fallback_content(topic, num_slides, focus_areas, custom_content)
    
    def _generate_with_gemini(self, topic: str, num_slides: int, focus_areas: List[str], 
                            custom_content: str, target_audience: str, presentation_style: str) -> Dict[str, Any]:
        """Generate content using Gemini AI"""
        
        prompt = f"""
        Create a comprehensive presentation structure for the topic: "{topic}"
        
        Requirements:
        - Number of content slides: {num_slides}
        - Target audience: {target_audience or 'General business audience'}
        - Presentation style: {presentation_style}
        - Focus areas: {', '.join(focus_areas) if focus_areas else 'balanced coverage'}
        
        {f"Additional custom content to incorporate: {custom_content}" if custom_content else ""}
        
        Please provide a JSON structure with:
        1. A compelling title and subtitle
        2. {num_slides} detailed content slides with:
           - Engaging titles
           - Detailed content descriptions (2-3 sentences)
           - 4-6 bullet points per slide
           - Image suggestions for each slide
        
        Focus on creating engaging, informative content that flows logically.
        Make the content domain-specific and relevant to the topic.
        
        Return ONLY valid JSON in this exact format:
        {{
            "title": "presentation title",
            "subtitle": "presentation subtitle", 
            "slides": [
                {{
                    "slide_number": 1,
                    "type": "content",
                    "title": "slide title",
                    "content": "detailed description",
                    "bullet_points": ["point1", "point2", "point3", "point4"],
                    "image_prompt": "description for image generation"
                }}
            ]
        }}
        """
        
        try:
            response = self.model.generate_content(prompt)
            response_text = response.text
            
            # Clean the response text
            response_text = response_text.strip()
            
            # Find JSON content between ```json and ``` or just extract JSON
            json_match = re.search(r'```(?:json)?\s*(.*?)\s*```', response_text, re.DOTALL)
            if json_match:
                json_content = json_match.group(1)
            else:
                # Try to find JSON-like content
                json_start = response_text.find('{')
                json_end = response_text.rfind('}')
                if json_start != -1 and json_end != -1:
                    json_content = response_text[json_start:json_end+1]
                else:
                    json_content = response_text
            
            # Parse JSON
            structure = json.loads(json_content)
            
            # Validate structure
            if not isinstance(structure, dict) or 'slides' not in structure:
                raise ValueError("Invalid JSON structure")
            
            # Ensure all slides have required fields
            for slide in structure['slides']:
                if 'bullet_points' not in slide:
                    slide['bullet_points'] = ["Key point 1", "Key point 2", "Key point 3"]
                if 'image_prompt' not in slide:
                    slide['image_prompt'] = f"{topic} related illustration"
                if 'content' not in slide:
                    slide['content'] = f"Content about {slide.get('title', topic)}"
            
            return structure
            
        except json.JSONDecodeError as e:
            st.warning(f"⚠️ Could not parse AI response as JSON: {str(e)}. Using fallback content generation.")
            return self._generate_fallback_content(topic, num_slides, focus_areas, custom_content)
        except Exception as e:
            st.warning(f"⚠️ Gemini AI error: {str(e)}. Using fallback content generation.")
            return self._generate_fallback_content(topic, num_slides, focus_areas, custom_content)
    
    def _generate_fallback_content(self, topic: str, num_slides: int, focus_areas: List[str], custom_content: str) -> Dict[str, Any]:
        """Fallback content generation when AI is not available"""
        
        structure = {
            "title": f"{topic}",
            "subtitle": "AI-Generated Presentation",
            "slides": []
        }
        
        # Incorporate custom content if provided
        if custom_content:
            custom_slide = {
                "slide_number": 1,
                "type": "custom",
                "title": f"About {topic}",
                "content": custom_content,
                "bullet_points": self._extract_key_points_from_text(custom_content),
                "image_prompt": f"{topic} overview illustration"
            }
            structure["slides"].append(custom_slide)
            num_slides -= 1
        
        # Define slide types based on focus areas
        if focus_areas:
            slide_types = focus_areas
        else:
            slide_types = ["introduction", "overview", "benefits", "challenges", "implementation", "conclusion"]
        
        for i in range(min(num_slides, len(slide_types))):
            slide_type = slide_types[i] if i < len(slide_types) else "detailed"
            slide_content = self._generate_slide_content(topic, slide_type)
            slide_content["slide_number"] = len(structure["slides"]) + 1
            structure["slides"].append(slide_content)
        
        return structure
    
    def _extract_key_points_from_text(self, text: str) -> List[str]:
        """Extract key points from custom content"""
        sentences = text.split('.')
        points = []
        for sentence in sentences[:4]:
            sentence = sentence.strip()
            if len(sentence) > 10:
                points.append(sentence)
        
        if len(points) < 3:
            points.extend([
                "Key insight from the provided content",
                "Important consideration for implementation", 
                "Strategic implications and next steps"
            ])
        
        return points[:6]
    
    def _generate_slide_content(self, topic: str, slide_type: str) -> Dict[str, Any]:
        """Generate content for a specific slide type with image prompts"""
        
        content_templates = {
            "introduction": {
                "title": f"Introduction to {topic}",
                "content": f"Welcome to our comprehensive presentation on {topic}. Today we'll explore the key aspects, benefits, and implications of this important subject.",
                "type": slide_type,
                "bullet_points": [
                    f"Understanding {topic}",
                    "Key concepts and definitions",
                    "Why this matters today",
                    "What we'll cover in this presentation"
                ],
                "image_prompt": f"{topic} introduction concept, modern professional illustration"
            },
            "overview": {
                "title": f"{topic} - Overview",
                "content": f"Let's examine the main components and aspects of {topic}.",
                "type": slide_type,
                "bullet_points": [
                    "Historical context and background",
                    "Current state and trends",
                    "Key stakeholders involved",
                    "Main applications and use cases"
                ],
                "image_prompt": f"{topic} overview infographic, business concept visualization"
            },
            "benefits": {
                "title": f"Benefits of {topic}",
                "content": f"Exploring the key advantages and positive impacts of {topic}.",
                "type": slide_type,
                "bullet_points": [
                    "Improved efficiency and productivity",
                    "Cost savings and resource optimization",
                    "Enhanced user experience",
                    "Future growth opportunities"
                ],
                "image_prompt": f"{topic} benefits illustration, growth and success concept"
            },
            "challenges": {
                "title": f"Challenges in {topic}",
                "content": f"Understanding the obstacles and considerations for {topic}.",
                "type": slide_type,
                "bullet_points": [
                    "Technical limitations and constraints",
                    "Implementation barriers",
                    "Resource requirements",
                    "Risk mitigation strategies"
                ],
                "image_prompt": f"{topic} challenges visualization, problem-solving concept"
            }
        }
        
        return content_templates.get(slide_type, {
            "title": f"{slide_type.replace('_', ' ').title()} - {topic}",
            "content": f"Detailed information about {topic} related to {slide_type}",
            "type": slide_type,
            "bullet_points": [
                "Key point about this aspect",
                "Important consideration",
                "Practical implementation",
                "Summary and takeaways"
            ],
            "image_prompt": f"{topic} {slide_type} professional illustration"
        })

class ImageGenerator:
    """Handles domain-related image generation using OpenAI DALL-E"""
    
    def __init__(self, api_key: str = None):
        self.api_key = api_key or OPENAI_API_KEY
        self.base_url = "https://api.openai.com/v1/images/generations"
        
    def generate_image(self, prompt: str, size: str = "1024x1024") -> Optional[bytes]:
        """Generate an image based on the prompt"""
        
        if not self.api_key:
            st.warning("⚠️ No OpenAI API key provided. Using placeholder image.")
            return self._generate_placeholder_image(prompt)
        
        try:
            headers = {
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json"
            }
            
            # Enhanced prompt for better professional images
            enhanced_prompt = f"Professional business presentation style: {prompt}. Clean, modern, suitable for corporate presentation. High quality, minimalist design, business appropriate."
            
            payload = {
                "model": "dall-e-3",
                "prompt": enhanced_prompt,
                "n": 1,
                "size": size,
                "quality": "standard"
            }
            
            st.info(f"🖼️ Generating image: {prompt[:50]}...")
            
            response = requests.post(self.base_url, headers=headers, json=payload, timeout=60)
            
            if response.status_code == 200:
                result = response.json()
                image_url = result["data"][0]["url"]
                
                # Download the image
                st.info("📥 Downloading generated image...")
                img_response = requests.get(image_url, timeout=30)
                if img_response.status_code == 200:
                    st.success("✅ Image generated successfully!")
                    return img_response.content
                else:
                    st.warning(f"⚠️ Failed to download image: {img_response.status_code}")
            else:
                error_msg = response.json().get('error', {}).get('message', 'Unknown error')
                st.error(f"❌ Image generation failed: {error_msg}")
                
        except requests.exceptions.Timeout:
            st.warning("⚠️ Image generation timeout. Using placeholder.")
        except requests.exceptions.RequestException as e:
            st.warning(f"⚠️ Network error during image generation: {str(e)}")
        except Exception as e:
            st.warning(f"⚠️ Image generation failed: {str(e)}")
        
        return self._generate_placeholder_image(prompt)
    
    def _generate_placeholder_image(self, prompt: str) -> bytes:
        """Generate a colored placeholder image"""
        try:
            # Create a professional looking placeholder
            img = Image.new('RGB', (1024, 768), color=(240, 248, 255))  # Light blue background
            draw = ImageDraw.Draw(img)
            
            # Add gradient effect
            for y in range(768):
                color_val = int(240 + (y / 768) * 15)  # Subtle gradient
                draw.line([(0, y), (1024, y)], fill=(color_val, 248, 255))
            
            # Add border
            draw.rectangle([(10, 10), (1014, 758)], outline=(100, 149, 237), width=5)
            
            # Add text
            try:
                font = ImageFont.load_default()
            except:
                font = None
                
            # Title text
            title_text = "Professional Image Placeholder"
            draw.text((512, 300), title_text, fill=(64, 64, 64), font=font, anchor="mm")
            
            # Description text
            desc_text = f"Topic: {prompt[:40]}..."
            draw.text((512, 350), desc_text, fill=(100, 100, 100), font=font, anchor="mm")
            
            # Additional info
            info_text = "Generated by AI PowerPoint Pro"
            draw.text((512, 400), info_text, fill=(150, 150, 150), font=font, anchor="mm")
            
            # Convert to bytes
            img_bytes = io.BytesIO()
            img.save(img_bytes, format='PNG', quality=95)
            return img_bytes.getvalue()
            
        except Exception as e:
            st.warning(f"⚠️ Error creating placeholder image: {str(e)}")
            # Return minimal placeholder
            img = Image.new('RGB', (800, 600), color=(200, 200, 200))
            img_bytes = io.BytesIO()
            img.save(img_bytes, format='PNG')
            return img_bytes.getvalue()

class EnhancedPresentationDesigner:
    """Enhanced presentation designer with multiple themes"""
    
    def __init__(self):
        self.themes = {
            "professional": {
                "name": "Professional",
                "description": "Clean, corporate style with blue accents",
                "background_color": RGBColor(255, 255, 255),
                "title_color": RGBColor(0, 51, 102),
                "content_color": RGBColor(64, 64, 64),
                "accent_color": RGBColor(0, 123, 191)
            },
            "modern": {
                "name": "Modern", 
                "description": "Contemporary design with subtle colors",
                "background_color": RGBColor(248, 249, 250),
                "title_color": RGBColor(33, 37, 41),
                "content_color": RGBColor(73, 80, 87),
                "accent_color": RGBColor(0, 123, 255)
            },
            "creative": {
                "name": "Creative",
                "description": "Vibrant and eye-catching design",
                "background_color": RGBColor(255, 255, 255),
                "title_color": RGBColor(138, 43, 226),
                "content_color": RGBColor(75, 0, 130),
                "accent_color": RGBColor(255, 140, 0)
            },
            "dark": {
                "name": "Dark",
                "description": "Modern dark theme",
                "background_color": RGBColor(45, 45, 45),
                "title_color": RGBColor(255, 255, 255),
                "content_color": RGBColor(220, 220, 220),
                "accent_color": RGBColor(100, 149, 237)
            }
        }
        
        self.current_theme = "professional"
    
    def set_theme(self, theme_name: str):
        if theme_name in self.themes:
            self.current_theme = theme_name
    
    def get_theme_colors(self) -> Dict[str, RGBColor]:
        return {k: v for k, v in self.themes[self.current_theme].items() if isinstance(v, RGBColor)}

class EnhancedPowerPointGenerator:
    """Enhanced PowerPoint generator with AI content and image generation"""
    
    def __init__(self, gemini_api_key: str = None, image_api_key: str = None):
        self.content_generator = GeminiContentGenerator(gemini_api_key)
        self.image_generator = ImageGenerator(image_api_key)
        self.designer = EnhancedPresentationDesigner()
    
    def create_presentation(self, topic: str, num_slides: int = 5, theme: str = "professional", 
                          focus_areas: List[str] = None, custom_content: str = None,
                          target_audience: str = None, presentation_style: str = "professional",
                          include_images: bool = False) -> bytes:
        """Create an enhanced PowerPoint presentation with AI content and images"""
        
        try:
            self.designer.set_theme(theme)
            
            # Generate content structure
            with st.spinner(f"🧠 Generating AI content for: {topic}..."):
                content_structure = self.content_generator.generate_content_structure(
                    topic, num_slides, focus_areas, custom_content, target_audience, presentation_style
                )
            
            st.success("✅ Content structure generated!")
            
            # Create presentation
            with st.spinner("🎨 Creating PowerPoint slides..."):
                prs = Presentation()
                
                # Add title slide
                self._create_title_slide(prs, content_structure["title"], content_structure["subtitle"])
                st.info("📄 Title slide created")
                
                # Add content slides with images
                for i, slide_data in enumerate(content_structure["slides"]):
                    slide_num = i + 1
                    st.info(f"📄 Creating slide {slide_num}: {slide_data.get('title', 'Untitled')}")
                    
                    if include_images:
                        try:
                            image_bytes = self.image_generator.generate_image(
                                slide_data.get("image_prompt", f"{topic} slide {slide_num}")
                            )
                            slide_data["image_bytes"] = image_bytes
                        except Exception as e:
                            st.warning(f"⚠️ Image generation failed for slide {slide_num}: {str(e)}")
                            slide_data["image_bytes"] = None
                    
                    self._create_content_slide_with_image(prs, slide_data, include_images)
                
                # Add closing slide
                self._create_closing_slide(prs, topic)
                st.info("📄 Closing slide created")
            
            st.success("✅ All slides created successfully!")
            
            # Convert to bytes
            with st.spinner("💾 Finalizing presentation..."):
                ppt_bytes = io.BytesIO()
                prs.save(ppt_bytes)
                ppt_bytes.seek(0)
                
            st.success("🎉 Presentation ready for download!")
            return ppt_bytes.getvalue()
            
        except Exception as e:
            st.error(f"❌ Error creating presentation: {str(e)}")
            raise e
    
    def _create_title_slide(self, prs, title: str, subtitle: str):
        try:
            slide_layout = prs.slide_layouts[0]
            slide = prs.slides.add_slide(slide_layout)
            
            title_shape = slide.shapes.title
            title_shape.text = title
            
            title_paragraph = title_shape.text_frame.paragraphs[0]
            title_paragraph.font.size = Pt(44)
            title_paragraph.font.color.rgb = self.designer.get_theme_colors()["title_color"]
            title_paragraph.alignment = PP_ALIGN.CENTER
            
            subtitle_shape = slide.placeholders[1]
            subtitle_shape.text = f"{subtitle}\nGenerated on {datetime.now().strftime('%B %d, %Y')}"
            
            for paragraph in subtitle_shape.text_frame.paragraphs:
                paragraph.font.size = Pt(24)
                paragraph.font.color.rgb = self.designer.get_theme_colors()["content_color"]
                paragraph.alignment = PP_ALIGN.CENTER
                
        except Exception as e:
            st.warning(f"⚠️ Error creating title slide: {str(e)}")
    
    def _create_content_slide_with_image(self, prs, slide_data: Dict[str, Any], include_images: bool = False):
        try:
            # Use a layout that supports images
            slide_layout = prs.slide_layouts[5] if include_images else prs.slide_layouts[1] 
            slide = prs.slides.add_slide(slide_layout)
            
            # Add title
            if include_images:
                # Create title manually for blank layout
                title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1))
                title_frame = title_shape.text_frame
                title_frame.text = slide_data["title"]
                title_paragraph = title_frame.paragraphs[0]
                title_paragraph.font.size = Pt(32)
                title_paragraph.font.color.rgb = self.designer.get_theme_colors()["title_color"]
                title_paragraph.alignment = PP_ALIGN.LEFT
                
                # Add image if available
                if slide_data.get("image_bytes"):
                    try:
                        image_stream = io.BytesIO(slide_data["image_bytes"])
                        slide.shapes.add_picture(image_stream, Inches(5.5), Inches(1.5), Inches(4), Inches(3))
                    except Exception as e:
                        st.warning(f"⚠️ Could not add image to slide: {str(e)}")
                
                # Add content text
                content_shape = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(4.5), Inches(5))
                text_frame = content_shape.text_frame
                
            else:
                # Use standard layout
                title_shape = slide.shapes.title
                title_shape.text = slide_data["title"]
                
                title_paragraph = title_shape.text_frame.paragraphs[0]
                title_paragraph.font.size = Pt(36)
                title_paragraph.font.color.rgb = self.designer.get_theme_colors()["title_color"]
                
                content_shape = slide.placeholders[1]
                text_frame = content_shape.text_frame
            
            text_frame.clear()
            
            # Add content
            if slide_data.get("content"):
                p = text_frame.paragraphs[0]
                p.text = slide_data["content"]
                p.font.size = Pt(16 if include_images else 18)
                p.font.color.rgb = self.designer.get_theme_colors()["content_color"]
                p.space_after = Pt(12)
            
            # Add bullet points
            for bullet_point in slide_data.get("bullet_points", []):
                p = text_frame.add_paragraph()
                p.text = bullet_point
                p.level = 1
                p.font.size = Pt(14 if include_images else 16)
                p.font.color.rgb = self.designer.get_theme_colors()["content_color"]
                p.space_after = Pt(6)
                
        except Exception as e:
            st.warning(f"⚠️ Error creating content slide: {str(e)}")
    
    def _create_closing_slide(self, prs, topic: str):
        try:
            slide_layout = prs.slide_layouts[5]
            slide = prs.slides.add_slide(slide_layout)
            
            title_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(1.5))
            title_frame = title_box.text_frame
            title_frame.text = "Thank You!"
            
            title_paragraph = title_frame.paragraphs[0]
            title_paragraph.font.size = Pt(48)
            title_paragraph.font.color.rgb = self.designer.get_theme_colors()["title_color"]
            title_paragraph.alignment = PP_ALIGN.CENTER
            
            subtitle_box = slide.shapes.add_textbox(Inches(1), Inches(4), Inches(8), Inches(2))
            subtitle_frame = subtitle_box.text_frame
            subtitle_frame.text = f"Questions & Discussion\n\nPresentation on: {topic}"
            
            for paragraph in subtitle_frame.paragraphs:
                paragraph.font.size = Pt(24)
                paragraph.font.color.rgb = self.designer.get_theme_colors()["content_color"]
                paragraph.alignment = PP_ALIGN.CENTER
                
        except Exception as e:
            st.warning(f"⚠️ Error creating closing slide: {str(e)}")

# Enhanced Streamlit App
def main():
    # Header
    st.title("🚀 AI PowerPoint Generator Pro")
    st.markdown("**Generate professional presentations with AI content and domain-specific images!**")
    
    # Show API status
    col1, col2 = st.columns(2)
    with col1:
        if GEMINI_AVAILABLE and GEMINI_API_KEY:
            st.success("🤖 Gemini AI: Ready")
        else:
            st.warning("⚠️ Gemini AI: Using fallback")
    
    with col2:
        if OPENAI_API_KEY:
            st.success("🖼️ Image Generation: Ready")
        else:
            st.warning("⚠️ Image Generation: Disabled")
    
    # Sidebar
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        # Topic input
        topic = st.text_input(
            "📝 Presentation Topic",
            value="Artificial Intelligence in Healthcare",
            help="Enter the main topic for your presentation"
        )
        
        # Target audience
        target_audience = st.selectbox(
            "🎯 Target Audience",
            ["General Business", "Technical Team", "Executives", "Students", "Investors", "General Public"],
            help="Select your target audience for tailored content"
        )
        
        # Presentation style
        presentation_style = st.selectbox(
            "🎨 Presentation Style",
            ["Professional", "Creative", "Educational", "Sales Pitch", "Technical Deep-dive"],
            help="Choose the overall tone and style"
        )
        
        # Number of slides
        num_slides = st.slider(
            "📊 Number of Content Slides",
            min_value=3,
            max_value=12, 
            value=6,
            help="Choose how many content slides you want"
        )
        
        # Theme selection
        st.subheader("🎨 Choose Theme")
        generator = EnhancedPowerPointGenerator()
        
        themes = generator.designer.themes
        theme_names = list(themes.keys())
        theme_descriptions = [themes[name]["description"] for name in theme_names]
        
        selected_theme = st.selectbox(
            "Select Theme",
            theme_names,
            format_func=lambda x: f"{themes[x]['name']} - {themes[x]['description']}",
            help="Choose your preferred theme"
        )
        
        # Focus areas
        st.subheader("🎯 Focus Areas")
        focus_options = [
            "Introduction & Overview",
            "Benefits & Advantages", 
            "Challenges & Solutions",
            "Implementation Strategy",
            "Case Studies & Examples",
            "Future Trends",
            "Technical Details",
            "Market Analysis",
            "Best Practices",
            "ROI & Metrics"
        ]
        
        selected_focus = st.multiselect(
            "Select focus areas (optional)",
            focus_options,
            help="Choose specific areas to emphasize in your presentation"
        )
        
        # Image generation toggle
        include_images = st.checkbox(
            "🖼️ Include AI-Generated Images",
            value=True,
            help="Generate domain-specific images for each slide (requires OpenAI API)"
        )
        
        # Custom content
        st.subheader("✍️ Custom Content")
        custom_content = st.text_area(
            "Additional Information",
            placeholder="Paste any specific content, data, or requirements you want to include...",
            height=100,
            help="Any additional content you want to incorporate into the presentation"
        )
    
    # Main content area
    if not topic.strip():
        st.warning("⚠️ Please enter a presentation topic to get started.")
        return
    
    # Preview section
    st.header("📋 Presentation Preview")
    
    col1, col2, col3 = st.columns([2, 1, 1])
    
    with col1:
        st.markdown(f"**Topic:** {topic}")
        st.markdown(f"**Audience:** {target_audience}")
        st.markdown(f"**Style:** {presentation_style}")
    
    with col2:
        st.markdown(f"**Slides:** {num_slides + 2} total")
        st.markdown(f"**Theme:** {themes[selected_theme]['name']}")
    
    with col3:
        st.markdown(f"**Images:** {'✅ Enabled' if include_images else '❌ Disabled'}")
        st.markdown(f"**Focus Areas:** {len(selected_focus)}")
    
    if selected_focus:
        st.markdown("**Selected Focus Areas:**")
        for area in selected_focus:
            st.markdown(f"• {area}")
    
    # Custom content preview
    if custom_content:
        with st.expander("📝 Custom Content Preview"):
            st.markdown(f'<div class="custom-content-box">{custom_content}</div>', 
                       unsafe_allow_html=True)
    
    # Generation section
    st.header("🚀 Generate Presentation")
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        if st.button("🎯 Generate AI PowerPoint", type="primary"):
            if not topic.strip():
                st.error("❌ Please enter a valid topic!")
                return
            
            try:
                # Create generator instance
                generator = EnhancedPowerPointGenerator(GEMINI_API_KEY, OPENAI_API_KEY)
                
                # Generate presentation
                start_time = time.time()
                
                ppt_bytes = generator.create_presentation(
                    topic=topic,
                    num_slides=num_slides,
                    theme=selected_theme,
                    focus_areas=selected_focus if selected_focus else None,
                    custom_content=custom_content if custom_content.strip() else None,
                    target_audience=target_audience,
                    presentation_style=presentation_style.lower().replace(" ", "_"),
                    include_images=include_images
                )
                
                generation_time = time.time() - start_time
                
                # Success message
                st.markdown(f'''
                <div class="success-message">
                    <h3>🎉 Presentation Generated Successfully!</h3>
                    <p><strong>Generation Time:</strong> {generation_time:.1f} seconds</p>
                    <p><strong>Total Slides:</strong> {num_slides + 2}</p>
                    <p><strong>Theme:</strong> {themes[selected_theme]["name"]}</p>
                    <p><strong>File Size:</strong> ~{len(ppt_bytes) / 1024:.1f} KB</p>
                </div>
                ''', unsafe_allow_html=True)
                
                # Download button
                filename = f"{topic.replace(' ', '_')}_presentation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
                
                st.download_button(
                    label="📥 Download PowerPoint",
                    data=ppt_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    help="Click to download your generated presentation"
                )
                
                # Additional info
                st.info(f"""
                📌 **Tips for your presentation:**
                • Review and customize the content as needed
                • Add your own branding and logos  
                • Practice your delivery with the generated content
                • Use the bullet points as speaking notes
                """)
                
            except Exception as e:
                st.error(f"❌ Generation failed: {str(e)}")
                st.info("💡 Try reducing the number of slides or disabling image generation if you continue to have issues.")
    
    with col2:
        st.markdown("### 💡 Pro Tips")
        st.markdown("""
        • Be specific with your topic
        • Select relevant focus areas
        • Include custom content for better results
        • Images enhance visual appeal
        • Review content before presenting
        """)
    
    # Footer information
    st.markdown("---")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("### 🤖 AI Features")
        st.markdown("""
        • Gemini AI content generation
        • DALL-E image creation
        • Smart topic analysis
        • Audience-tailored content
        """)
    
    with col2:
        st.markdown("### 🎨 Design Features")
        st.markdown("""
        • Multiple professional themes
        • Consistent formatting
        • Optimized layouts
        • Visual hierarchy
        """)
    
    with col3:
        st.markdown("### 📈 Export Features")
        st.markdown("""
        • Standard .pptx format
        • Compatible with PowerPoint
        • Editable content
        • High-quality images
        """)
    
    # Usage statistics (if you want to track)
    if 'generation_count' not in st.session_state:
        st.session_state.generation_count = 0
    
    st.markdown(f"""
    <div class="info-box">
        <small>
        🔧 <strong>System Status:</strong> All systems operational<br>
        📊 <strong>Session Generations:</strong> {st.session_state.generation_count}<br>
        ⏰ <strong>Last Updated:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
        </small>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
