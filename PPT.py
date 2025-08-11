import streamlit as st
import os
import json
import re
from typing import List, Dict, Any
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from datetime import datetime
import io
import base64

# Set page config
st.set_page_config(
    page_title="AI PowerPoint Generator",
    page_icon="📊",
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
</style>
""", unsafe_allow_html=True)

class AIContentGenerator:
    """Handles AI-powered content generation for presentations"""
    
    def __init__(self, api_key: str = None):
        self.api_key = api_key
    
    def generate_content_structure(self, topic: str, num_slides: int = 5, focus_areas: List[str] = None) -> Dict[str, Any]:
        """Generate a structured presentation outline"""
        
        structure = {
            "title": f"{topic}",
            "subtitle": "AI-Generated Presentation",
            "slides": []
        }
        
        # Define slide types based on user focus areas or defaults
        if focus_areas:
            slide_types = focus_areas
        else:
            slide_types = ["introduction", "overview", "detailed", "benefits", "challenges", "conclusion"]
        
        for i in range(min(num_slides, len(slide_types))):
            slide_type = slide_types[i] if i < len(slide_types) else "detailed"
            slide_content = self._generate_slide_content(topic, slide_type)
            
            structure["slides"].append({
                "slide_number": i + 1,
                "type": slide_type,
                "title": slide_content["title"],
                "content": slide_content["content"],
                "bullet_points": slide_content.get("bullet_points", [])
            })
        
        return structure
    
    def _generate_slide_content(self, topic: str, slide_type: str) -> Dict[str, Any]:
        """Generate content for a specific slide type"""
        
        content_templates = {
            "introduction": {
                "title": f"Introduction to {topic}",
                "content": f"Welcome to our comprehensive presentation on {topic}. Today we'll explore the key aspects, benefits, and implications of this important subject.",
                "bullet_points": [
                    f"Understanding {topic}",
                    "Key concepts and definitions",
                    "Why this matters today",
                    "What we'll cover in this presentation"
                ]
            },
            "overview": {
                "title": f"{topic} - Overview",
                "content": f"Let's examine the main components and aspects of {topic}.",
                "bullet_points": [
                    "Historical context and background",
                    "Current state and trends",
                    "Key stakeholders involved",
                    "Main applications and use cases"
                ]
            },
            "detailed": {
                "title": f"Deep Dive into {topic}",
                "content": f"Here's a detailed analysis of the core elements of {topic}.",
                "bullet_points": [
                    "Technical specifications and requirements",
                    "Implementation strategies",
                    "Best practices and methodologies",
                    "Real-world examples and case studies"
                ]
            },
            "benefits": {
                "title": f"Benefits of {topic}",
                "content": f"Exploring the key advantages and positive impacts of {topic}.",
                "bullet_points": [
                    "Improved efficiency and productivity",
                    "Cost savings and resource optimization",
                    "Enhanced user experience",
                    "Future growth opportunities"
                ]
            },
            "challenges": {
                "title": f"Challenges in {topic}",
                "content": f"Understanding the obstacles and considerations for {topic}.",
                "bullet_points": [
                    "Technical limitations and constraints",
                    "Implementation barriers",
                    "Resource requirements",
                    "Risk mitigation strategies"
                ]
            },
            "conclusion": {
                "title": f"Conclusion - {topic}",
                "content": f"Summary of key takeaways from our discussion on {topic}.",
                "bullet_points": [
                    "Recap of main points covered",
                    "Strategic recommendations",
                    "Next steps and action items",
                    "Questions and discussion"
                ]
            },
            "market_analysis": {
                "title": f"Market Analysis - {topic}",
                "content": f"Current market trends and opportunities in {topic}.",
                "bullet_points": [
                    "Market size and growth projections",
                    "Key players and competitors",
                    "Market opportunities and gaps",
                    "Consumer behavior and preferences"
                ]
            },
            "technical_specs": {
                "title": f"Technical Specifications - {topic}",
                "content": f"Technical details and specifications for {topic}.",
                "bullet_points": [
                    "System requirements and architecture",
                    "Performance metrics and benchmarks",
                    "Security and compliance standards",
                    "Integration capabilities"
                ]
            }
        }
        
        return content_templates.get(slide_type, {
            "title": f"{slide_type.replace('_', ' ').title()} - {topic}",
            "content": f"Detailed information about {topic}",
            "bullet_points": [
                "Key point 1",
                "Key point 2", 
                "Key point 3",
                "Summary and takeaways"
            ]
        })

class PresentationDesigner:
    """Handles presentation design and formatting"""
    
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

class StreamlitPowerPointGenerator:
    """Streamlit-optimized PowerPoint generator"""
    
    def __init__(self):
        self.content_generator = AIContentGenerator()
        self.designer = PresentationDesigner()
    
    def create_presentation(self, topic: str, num_slides: int = 5, theme: str = "professional", focus_areas: List[str] = None) -> bytes:
        """Create a PowerPoint presentation and return as bytes"""
        
        self.designer.set_theme(theme)
        
        # Generate content structure
        with st.spinner(f"🤖 Generating AI content for: {topic}..."):
            content_structure = self.content_generator.generate_content_structure(topic, num_slides, focus_areas)
        
        # Create presentation
        with st.spinner("📊 Creating PowerPoint slides..."):
            prs = Presentation()
            
            # Add title slide
            self._create_title_slide(prs, content_structure["title"], content_structure["subtitle"])
            
            # Add content slides
            for slide_data in content_structure["slides"]:
                self._create_content_slide(prs, slide_data)
            
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
    
    def _create_content_slide(self, prs, slide_data: Dict[str, Any]):
        slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(slide_layout)
        
        title_shape = slide.shapes.title
        title_shape.text = slide_data["title"]
        
        title_paragraph = title_shape.text_frame.paragraphs[0]
        title_paragraph.font.size = Pt(36)
        title_paragraph.font.color.rgb = self.designer.get_theme_colors()["title_color"]
        
        content_shape = slide.placeholders[1]
        text_frame = content_shape.text_frame
        text_frame.clear()
        
        if slide_data.get("content"):
            p = text_frame.paragraphs[0]
            p.text = slide_data["content"]
            p.font.size = Pt(18)
            p.font.color.rgb = self.designer.get_theme_colors()["content_color"]
            p.space_after = Pt(12)
        
        for bullet_point in slide_data.get("bullet_points", []):
            p = text_frame.add_paragraph()
            p.text = bullet_point
            p.level = 1
            p.font.size = Pt(20)
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

# Streamlit App
def main():
    # Header
    st.title("🤖 AI PowerPoint Generator")
    st.markdown("**Generate professional presentations with AI in seconds!**")
    
    # Sidebar
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        # Topic input
        topic = st.text_input(
            "📝 Presentation Topic", 
            value="Artificial Intelligence in Business",
            help="Enter the main topic for your presentation"
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
        generator = StreamlitPowerPointGenerator()
        
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
        st.subheader("🎯 Focus Areas (Optional)")
        focus_areas = st.multiselect(
            "Select specific areas to focus on:",
            ["introduction", "overview", "benefits", "challenges", "market_analysis", "technical_specs", "conclusion"],
            default=["introduction", "overview", "benefits", "challenges", "conclusion"],
            help="Choose which aspects to emphasize in your presentation"
        )
        
        # Advanced options
        with st.expander("🔧 Advanced Options"):
            include_charts = st.checkbox("Include Chart Placeholders", value=False)
            include_images = st.checkbox("Include Image Placeholders", value=False)
            custom_footer = st.text_input("Custom Footer Text", value="")
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("📋 Presentation Preview")
        
        if topic:
            # Preview the structure
            preview_structure = generator.content_generator.generate_content_structure(
                topic, num_slides, focus_areas if focus_areas else None
            )
            
            st.markdown("**Your presentation will include:**")
            
            # Title slide preview
            st.markdown("**1. Title Slide**")
            st.markdown(f"- **Title**: {preview_structure['title']}")
            st.markdown(f"- **Subtitle**: {preview_structure['subtitle']}")
            
            # Content slides preview
            for i, slide in enumerate(preview_structure['slides'], 2):
                with st.expander(f"**{i}. {slide['title']}**"):
                    st.markdown(f"**Content**: {slide['content']}")
                    st.markdown("**Key Points**:")
                    for point in slide['bullet_points']:
                        st.markdown(f"• {point}")
            
            # Closing slide
            st.markdown(f"**{len(preview_structure['slides']) + 2}. Thank You Slide**")
        
        else:
            st.info("👆 Enter a topic in the sidebar to see the presentation preview")
    
    with col2:
        st.subheader("🚀 Generate Presentation")
        
        if st.button("🎯 Generate PowerPoint", type="primary", disabled=not topic):
            if topic:
                try:
                    # Generate presentation
                    ppt_bytes = generator.create_presentation(
                        topic=topic,
                        num_slides=num_slides,
                        theme=selected_theme,
                        focus_areas=focus_areas if focus_areas else None
                    )
                    
                    # Success message
                    st.markdown('<div class="success-message">✅ <strong>Success!</strong> Your presentation has been generated!</div>', unsafe_allow_html=True)
                    
                    # Generate filename
                    clean_topic = re.sub(r'[^\w\s-]', '', topic)
                    clean_topic = re.sub(r'[-\s]+', '_', clean_topic)
                    filename = f"AI_Presentation_{clean_topic}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
                    
                    # Download button
                    st.download_button(
                        label="📥 Download PowerPoint",
                        data=ppt_bytes,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        type="primary"
                    )
                    
                    # Display file info
                    st.info(f"**File**: {filename}  \n**Size**: {len(ppt_bytes) / 1024:.1f} KB  \n**Slides**: {num_slides + 2}")
                    
                except Exception as e:
                    st.error(f"❌ Error generating presentation: {str(e)}")
                    st.info("💡 Make sure you have installed: `pip install python-pptx streamlit`")
        
        # Quick tips
        with st.expander("💡 Tips for Better Presentations"):
            st.markdown("""
            **For better results:**
            - Use specific, descriptive topics
            - Choose 5-8 slides for optimal length
            - Select focus areas relevant to your audience
            - Professional theme works best for business
            - Creative theme is great for educational content
            """)
        
        # Example topics
        with st.expander("📚 Example Topics"):
            example_topics = [
                "Digital Marketing Strategy 2024",
                "Sustainable Energy Solutions",
                "Remote Work Best Practices",
                "Machine Learning for Beginners",
                "Cybersecurity Fundamentals",
                "Project Management Methodologies",
                "Customer Experience Design",
                "Data Analytics in Healthcare"
            ]
            
            for example_topic in example_topics:
                if st.button(f"Use: {example_topic}", key=f"example_{example_topic}"):
                    st.session_state.topic = example_topic
                    st.rerun()

    # Footer
    st.markdown("---")
    st.markdown("Made with ❤️ using Streamlit and python-pptx | Generate professional presentations in seconds!")

if __name__ == "__main__":
    main()
