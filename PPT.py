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
from PIL import Image
import google.generativeai as genai

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
</style>
""", unsafe_allow_html=True)

class GeminiContentGenerator:
    """Enhanced AI content generator using Gemini AI"""
    
    def __init__(self, api_key: str = None):
        self.api_key = api_key
        if api_key:
            genai.configure(api_key=api_key)
            self.model = genai.GenerativeModel('gemini-pro')
        else:
            self.model = None
    
    def generate_content_structure(self, topic: str, num_slides: int = 5, 
                                 focus_areas: List[str] = None, 
                                 custom_content: str = None,
                                 target_audience: str = None,
                                 presentation_style: str = "professional") -> Dict[str, Any]:
        """Generate AI-powered presentation structure using Gemini"""
        
        if self.model and self.api_key:
            return self._generate_with_gemini(topic, num_slides, focus_areas, custom_content, target_audience, presentation_style)
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
        
        Return in JSON format:
        {{
            "title": "presentation title",
            "subtitle": "presentation subtitle", 
            "slides": [
                {{
                    "slide_number": 1,
                    "type": "slide_type",
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
            # Extract JSON from response
            response_text = response.text
            
            # Find JSON content between ```json and ```
            json_match = re.search(r'```(?:json)?\s*(.*?)\s*```', response_text, re.DOTALL)
            if json_match:
                json_content = json_match.group(1)
            else:
                # If no code blocks, try to find JSON-like content
                json_content = response_text
            
            # Parse JSON
            try:
                structure = json.loads(json_content)
                return structure
            except json.JSONDecodeError:
                st.warning("⚠️ Could not parse AI response as JSON. Using fallback content generation.")
                return self._generate_fallback_content(topic, num_slides, focus_areas, custom_content)
                
        except Exception as e:
            st.error(f"❌ Error with Gemini AI: {str(e)}")
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
            slide_types = ["introduction", "overview", "detailed", "benefits", "challenges", "conclusion"]
        
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
        for sentence in sentences[:4]:  # Take first 4 sentences
            sentence = sentence.strip()
            if len(sentence) > 10:
                points.append(sentence)
        
        if len(points) < 3:
            points.extend([
                "Key insight from the provided content",
                "Important consideration for implementation", 
                "Strategic implications and next steps"
            ])
        
        return points[:6]  # Limit to 6 points
    
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
            "content": f"Detailed information about {topic}",
            "type": slide_type,
            "bullet_points": [
                "Key point 1",
                "Key point 2",
                "Key point 3", 
                "Summary and takeaways"
            ],
            "image_prompt": f"{topic} {slide_type} professional illustration"
        })

class ImageGenerator:
    """Handles domain-related image generation"""
    
    def __init__(self, api_key: str = None):
        self.api_key = api_key
        self.base_url = "https://api.openai.com/v1/images/generations"  # OpenAI DALL-E
        # Alternative: You can use Stability AI, Midjourney API, etc.
        
    def generate_image(self, prompt: str, size: str = "1024x1024") -> Optional[bytes]:
        """Generate an image based on the prompt"""
        
        if not self.api_key:
            return self._generate_placeholder_image(prompt)
        
        try:
            headers = {
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json"
            }
            
            payload = {
                "model": "dall-e-3",
                "prompt": f"Professional business presentation style: {prompt}. Clean, modern, suitable for corporate presentation.",
                "n": 1,
                "size": size,
                "quality": "standard"
            }
            
            response = requests.post(self.base_url, headers=headers, json=payload, timeout=30)
            
            if response.status_code == 200:
                result = response.json()
                image_url = result["data"][0]["url"]
                
                # Download the image
                img_response = requests.get(image_url, timeout=30)
                if img_response.status_code == 200:
                    return img_response.content
            
        except Exception as e:
            st.warning(f"⚠️ Image generation failed: {str(e)}. Using placeholder.")
        
        return self._generate_placeholder_image(prompt)
    
    def _generate_placeholder_image(self, prompt: str) -> bytes:
        """Generate a colored placeholder image"""
        from PIL import Image, ImageDraw, ImageFont
        
        # Create a colored rectangle as placeholder
        img = Image.new('RGB', (800, 600), color=(100, 149, 237))
        draw = ImageDraw.Draw(img)
        
        # Add text
        try:
            font = ImageFont.load_default()
        except:
            font = None
            
        text = f"Image: {prompt[:50]}..."
        draw.text((50, 250), text, fill=(255, 255, 255), font=font)
        draw.text((50, 300), "[AI Image Placeholder]", fill=(255, 255, 255), font=font)
        
        # Convert to bytes
        img_bytes = io.BytesIO()
        img.save(img_bytes, format='PNG')
        return img_bytes.getvalue()

class EnhancedPresentationDesigner:
    """Enhanced presentation designer with image support"""
    
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
        
        self.designer.set_theme(theme)
        
        # Generate content structure
        with st.spinner(f"🧠 Generating AI content for: {topic}..."):
            content_structure = self.content_generator.generate_content_structure(
                topic, num_slides, focus_areas, custom_content, target_audience, presentation_style
            )
        
        # Create presentation
        with st.spinner("🎨 Creating PowerPoint slides..."):
            prs = Presentation()
            
            # Add title slide
            self._create_title_slide(prs, content_structure["title"], content_structure["subtitle"])
            
            # Add content slides with images
            for i, slide_data in enumerate(content_structure["slides"]):
                if include_images:
                    with st.spinner(f"🖼️ Generating image for slide {i+1}..."):
                        image_bytes = self.image_generator.generate_image(slide_data.get("image_prompt", f"{topic} slide {i+1}"))
                        slide_data["image_bytes"] = image_bytes
                
                self._create_content_slide_with_image(prs, slide_data, include_images)
            
            # Add closing slide
            self._create_closing_slide(prs, topic)
        
        # Convert to bytes
        ppt_bytes = io.BytesIO()
        prs.save(ppt_bytes)
        ppt_bytes.seek(0)
        
        return ppt_bytes.getvalue()
    
    def _create_title_slide(self, prs, title: str, subtitle: str):
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
    
    def _create_content_slide_with_image(self, prs, slide_data: Dict[str, Any], include_images: bool = False):
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
    
    def _create_closing_slide(self, prs, topic: str):
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

# Enhanced Streamlit App
def main():
    # Header
    st.title("🚀 AI PowerPoint Generator Pro")
    st.markdown("**Generate professional presentations with AI content and domain-specific images!**")
    
    # API Keys Configuration
    with st.expander("🔑 API Configuration", expanded=False):
        col1, col2 = st.columns(2)
        
        with col1:
            gemini_api_key = st.text_input(
                "Gemini API Key (for AI content)",
                type="password",
                help="Get your API key from Google AI Studio: https://makersuite.google.com/app/apikey"
            )
        
        with col2:
            image_api_key = st.text_input(
                "Image Generation API Key (OpenAI)",
                type="password", 
                help="Get your API key from OpenAI: https://platform.openai.com/api-keys"
            )
    
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
            max_value=15, 
            value=6,
            help="Choose how many content slides you want"
        )
        
        # Theme selection
        st.subheader("🎨 Choose Theme")
        generator = EnhancedPowerPointGenerator(gemini_api_key, image_api_key)
        
        theme_options = {}
        for theme_key, theme_data in generator.designer.themes.items():
            theme_options[theme_data["name"]] = theme_key
        
        selected_theme_name = st.selectbox(
            "Theme Style",
            options=list(theme_options.keys()),
            help="Select the visual style for your presentation"
        )
        selected_theme = theme_options[selected_theme_name]
        
        # Display theme info
        theme_info = generator.designer.themes[selected_theme]
        st.info(f"**{theme_info['name']}**: {theme_info['description']}")
        
        # Focus areas
        st.subheader("🎯 Focus Areas")
        focus_areas = st.multiselect(
            "Select specific areas to focus on:",
            ["introduction", "overview", "benefits", "challenges", "market_analysis", "technical_specs", "conclusion"],
            default=["introduction", "overview", "benefits", "challenges"],
            help="Choose which aspects to emphasize in your presentation"
        )
        
        # Image generation
        include_images = st.checkbox(
            "🖼️ Generate Domain-Related Images",
            value=bool(image_api_key),
            help="Generate relevant images for each slide (requires API key)"
        )
        
        if include_images and not image_api_key:
            st.warning("⚠️ Image generation requires an API key. Placeholder images will be used.")
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("📝 Custom Content Input")
        
        # Custom content input
        custom_content = st.text_area(
            "Add your custom content (optional):",
            height=200,
            placeholder="Enter any specific information, data, or points you want to include in your presentation. This will be incorporated along with AI-generated content.",
            help="This content will be intelligently integrated into your presentation"
        )
        
        if custom_content:
            st.markdown('<div class="custom-content-box">✅ <strong>Custom content will be incorporated</strong> into your presentation along with AI-generated content.</div>', unsafe_allow_html=True)
        
        # Preview section
        st.subheader("📋 Presentation Preview")
        
        if topic:
            # Show what will be generated
            st.markdown("**Your presentation will include:**")
            st.markdown(f"- **Topic**: {topic}")
            st.markdown(f"- **Target Audience**: {target_audience}")
            st.markdown(f"- **Style**: {presentation_style}")
            st.markdown(f"- **Content Slides**: {num_slides}")
            st.markdown(f"- **Theme**: {selected_theme_name}")
            st.markdown(f"- **Images**: {'✅ AI-generated' if include_images else '❌ Text only'}")
            st.markdown(f"- **Custom Content**: {'✅ Included' if custom_content else '❌ AI-only'}")
            
            if gemini_api_key:
                st.success("🤖 **Gemini AI** will generate intelligent, tailored content")
            else:
                st.info("💡 Add Gemini API key for enhanced AI content generation")
        
        else:
            st.info("👆 Enter a topic in the sidebar to see the presentation preview")
    
    with col2:
        st.subheader("🚀 Generate Presentation")
        
        # Generate button
        if st.button("🎯 Generate AI PowerPoint", type="primary", disabled=not topic):
            if topic:
                try:
                    # Create enhanced generator instance
                    generator = EnhancedPowerPointGenerator(gemini_api_key, image_api_key)
                    
                    # Generate presentation
                    ppt_bytes = generator.create_presentation(
                        topic=topic,
                        num_slides=num_slides,
                        theme=selected_theme,
                        focus_areas=focus_areas if focus_areas else None,
                        custom_content=custom_content if custom_content else None,
                        target_audience=target_audience,
                        presentation_style=presentation_style.lower(),
                        include_images=include_images
                    )
                    
                    # Success message
                    st.markdown('<div class="success-message">✅ <strong>Success!</strong> Your AI-powered presentation has been generated!</div>', unsafe_allow_html=True)
                    
                    # Generate filename
                    clean_topic = re.sub(r'[^\w\s-]', '', topic)
                    clean_topic = re.sub(r'[-\s]+', '_', clean_topic)
                    filename = f"AI_Pro_Presentation_{clean_topic}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
                    
                    # Download button
                    st.download_button(
                        label="📥 Download PowerPoint",
                        data=ppt_bytes,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        type="primary"
                    )
                    
                    # Display file info
                    features_used = []
                    if gemini_api_key:
                        features_used.append("Gemini AI Content")
                    if include_images:
                        features_used.append("AI Images")
                    if custom_content:
                        features_used.append("Custom Content")
                    
                    st.info(f"""
                    **File**: {filename}  
                    **Size**: {len(ppt_bytes) / 1024:.1f} KB  
                    **Slides**: {num_slides + 2}  
                    **Features**: {', '.join(features_used) if features_used else 'Standard Generation'}
                    """)
                    
                except Exception as e:
                    st.error(f"❌ Error generating presentation: {str(e)}")
                    st.info("💡 Check your API keys and internet connection.")
                    
            else:
                st.warning("⚠️ Please enter a topic to generate the presentation")
        
        # Additional options
        st.markdown("---")
        st.subheader("💡 Tips")
        st.markdown("""
        **For best results:**
        - Use specific, clear topics
        - Add custom content for personalization
        - Choose appropriate audience
        - Enable images for visual impact
        
        **API Keys needed for:**
        - 🤖 **Gemini AI**: Enhanced content generation
        - 🖼️ **OpenAI**: Professional images
        """)
        
        # Quick examples
        st.subheader("🎯 Topic Examples")
        example_topics = [
            "Digital Marketing Strategy 2025",
            "Sustainable Energy Solutions", 
            "Machine Learning in Finance",
            "Remote Work Best Practices",
            "Cybersecurity Fundamentals"
        ]
        
        for example in example_topics:
            if st.button(f"📝 {example}", key=example):
                # Use JavaScript to set the topic input
                st.rerun()
    
    # Footer information
    st.markdown("---")
    st.markdown("""
    <div class="info-box">
    <h4>🔧 How it Works</h4>
    <ol>
    <li><strong>AI Content Generation</strong>: Gemini AI creates tailored content based on your topic and preferences</li>
    <li><strong>Smart Slide Design</strong>: Automatically structures content into logical, flowing slides</li>
    <li><strong>Image Integration</strong>: Generates relevant, professional images for each slide</li>
    <li><strong>Theme Application</strong>: Applies consistent, professional styling throughout</li>
    <li><strong>Export Ready</strong>: Creates standard .pptx files compatible with PowerPoint and Google Slides</li>
    </ol>
    </div>
    """, unsafe_allow_html=True)
    
    # Usage statistics (if you want to track)
    with st.expander("📊 Feature Status", expanded=False):
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if gemini_api_key:
                st.success("✅ AI Content Generation")
            else:
                st.info("ℹ️ Using Fallback Content")
        
        with col2:
            if image_api_key and include_images:
                st.success("✅ AI Image Generation")
            else:
                st.info("ℹ️ Text-Only Mode")
        
        with col3:
            if custom_content:
                st.success("✅ Custom Content Added")
            else:
                st.info("ℹ️ AI-Only Content")

if __name__ == "__main__":
    main()
