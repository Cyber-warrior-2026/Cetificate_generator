import streamlit as st
import openpyxl
from PIL import Image, ImageDraw, ImageFont
import io
import base64
import zipfile
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import tempfile
import os

# Set page config
st.set_page_config(
    page_title="Certificate Generator",
    page_icon="🎓",
    layout="wide",
)

# Initialize session state variables
if 'text_elements' not in st.session_state:
    st.session_state.text_elements = []
if 'template_size' not in st.session_state:
    st.session_state.template_size = None
if 'excel_headers' not in st.session_state:
    st.session_state.excel_headers = []
if 'certificates_generated' not in st.session_state:
    st.session_state.certificates_generated = False
if 'certificate_files' not in st.session_state:
    st.session_state.certificate_files = {}

# Title and description
st.title("Certificate Generator")
st.markdown("""
Generate beautiful certificates from an Excel list of participants.
Upload your certificate template and Excel data, position your text fields, and download all certificates in one click!
""")

# Create tabs for workflow
tabs = st.tabs(["1. Upload Files", "2. Design Certificate", "3. Generate & Send"])

# Tab 1: Upload Files
with tabs[0]:
    st.header("Upload Your Files")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Upload certificate template
        st.subheader("Certificate Template")
        template_file = st.file_uploader("Upload certificate template image", type=["png", "jpg", "jpeg"])
        
        if template_file:
            # Display the template
            image = Image.open(template_file)
            st.session_state.template_size = image.size
            st.image(image, caption="Certificate Template", use_column_width=True)
            st.success(f"✅ Template loaded: {image.width}x{image.height} pixels")
    
    with col2:
        # Upload Excel file
        st.subheader("Participant Data")
        excel_file = st.file_uploader("Upload Excel file with participant data", type=["xlsx", "xls"])
        
        if excel_file:
            try:
                # Load Excel data
                wb = openpyxl.load_workbook(excel_file)
                sheet = wb.active
                headers = [str(cell.value) for cell in sheet[1] if cell.value]
                st.session_state.excel_headers = headers
                
                # Show Excel preview
                preview_data = []
                for i, row in enumerate(sheet.iter_rows(values_only=True)):
                    if i == 0:  # Headers
                        preview_data.append(row[:5])  # Show first 5 columns
                    elif i < 6:  # First 5 rows
                        preview_data.append(row[:5])
                    else:
                        break
                
                st.dataframe(preview_data, use_container_width=True)
                st.success(f"✅ Excel loaded: {sheet.max_row-1} participants, {len(headers)} columns")
                
                # Email configuration
                st.subheader("Email Configuration (Optional)")
                email_col = st.selectbox("Select email column", options=["None"] + headers)
                
                if email_col != "None":
                    st.session_state.email_column = email_col
                    col_a, col_b = st.columns(2)
                    with col_a:
                        st.session_state.sender_email = st.text_input("Your email address")
                    with col_b:
                        st.session_state.email_password = st.text_input("App password", type="password")
                    
                    st.info("To send emails, you'll need an App Password if using Gmail. [Learn how](https://support.google.com/accounts/answer/185833)")
            
            except Exception as e:
                st.error(f"Error loading Excel file: {str(e)}")
    
    # Navigation buttons
    if template_file and excel_file:
        if st.button("Continue to Design ➡️", use_container_width=True):
            # Switch to the next tab
            st.session_state.active_tab = 1

# Tab 2: Design Certificate
with tabs[1]:
    st.header("Design Your Certificate")
    
    # Check if files are loaded
    if not st.session_state.get('template_size') or not st.session_state.excel_headers:
        st.warning("⚠️ Please upload your template and Excel file in the previous tab first")
    else:
        col1, col2 = st.columns([1, 2])
        
        with col1:
            st.subheader("Add Text Fields")
            
            # Field selection
            field = st.selectbox("Select field from Excel", options=st.session_state.excel_headers)
            
            # Text properties
            font_size = st.slider("Font size", min_value=10, max_value=100, value=36)
            color = st.color_picker("Text color", "#000000")
            
            # Position with sliders
            st.subheader("Position")
            x_pos = st.slider("X position (%)", min_value=0, max_value=100, value=50)
            y_pos = st.slider("Y position (%)", min_value=0, max_value=100, value=50)
            
            # Add field button
            if st.button("Add Field", use_container_width=True):
                # Calculate actual coordinates
                actual_x = int(st.session_state.template_size[0] * x_pos / 100)
                actual_y = int(st.session_state.template_size[1] * y_pos / 100)
                
                # Add to session state
                st.session_state.text_elements.append({
                    'field': field,
                    'font_size': font_size,
                    'color': color,
                    'x_pos': x_pos,
                    'y_pos': y_pos,
                    'actual_x': actual_x,
                    'actual_y': actual_y
                })
                st.success(f"✅ Added field: {field}")
            
            # List current fields
            st.subheader("Current Fields")
            if not st.session_state.text_elements:
                st.info("No fields added yet")
            else:
                for i, element in enumerate(st.session_state.text_elements):
                    col_a, col_b = st.columns([3, 1])
                    with col_a:
                        st.write(f"{i+1}. **{element['field']}** (Size: {element['font_size']})")
                    with col_b:
                        if st.button("🗑️", key=f"del_{i}"):
                            st.session_state.text_elements.pop(i)
                            st.experimental_rerun()
            
            # Clear all button
            if st.session_state.text_elements:
                if st.button("Clear All Fields", use_container_width=True):
                    st.session_state.text_elements = []
                    st.experimental_rerun()
        
        with col2:
            st.subheader("Certificate Preview")
            
            # Load the template
            if 'template_file' not in st.session_state and template_file:
                st.session_state.template_file = template_file
            
            if hasattr(st.session_state, 'template_file') and st.session_state.template_file:
                # Create preview image
                image = Image.open(st.session_state.template_file)
                draw = ImageDraw.Draw(image)
                
                # Draw each text element
                for element in st.session_state.text_elements:
                    x = element['actual_x']
                    y = element['actual_y']
                    
                    # Use default font for preview
                    try:
                        font = ImageFont.truetype("Arial.ttf", element['font_size'])
                    except:
                        font = ImageFont.load_default()
                    
                    # Draw text
                    draw.text((x, y), element['field'], fill=element['color'], font=font, anchor="mm")
                
                # Show preview
                st.image(image, caption="Certificate Preview", use_column_width=True)
            else:
                st.info("Please upload a template image in the previous tab")
        
        # Navigation
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("⬅️ Back to Upload", use_container_width=True):
                st.session_state.active_tab = 0
        with col_b:
            if len(st.session_state.text_elements) > 0:
                if st.button("Continue to Generate ➡️", use_container_width=True):
                    st.session_state.active_tab = 2
            else:
                st.button("Add fields to continue", disabled=True, use_container_width=True)

# Tab 3: Generate & Send
with tabs[2]:
    st.header("Generate & Send Certificates")
    
    # Check if design is ready
    if not st.session_state.get('template_size') or not st.session_state.excel_headers or not st.session_state.text_elements:
        st.warning("⚠️ Please complete the previous steps first")
    else:
        # Generate button
        if st.button("🎓 GENERATE CERTIFICATES", type="primary", use_container_width=True):
            with st.spinner("Generating certificates..."):
                try:
                    # Load template and Excel
                    template_image = Image.open(st.session_state.template_file)
                    wb = openpyxl.load_workbook(excel_file)
                    sheet = wb.active
                    
                    # Create temporary directory
                    temp_dir = tempfile.mkdtemp()
                    
                    # Create a zip file in memory
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
                        # Generate each certificate
                        certificate_count = 0
                        cert_files = {}
                        
                        for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
                            # Skip empty rows
                            if not any(row):
                                continue
                                
                            # Create a copy of the template
                            certificate = template_image.copy()
                            draw = ImageDraw.Draw(certificate)
                            
                            # Add all fields
                            for element in st.session_state.text_elements:
                                field_name = element['field']
                                col_idx = st.session_state.excel_headers.index(field_name)
                                
                                # Get value from Excel
                                text = str(row[col_idx]) if row[col_idx] is not None else ""
                                
                                # Use best available font
                                try:
                                    font = ImageFont.truetype("Arial.ttf", element['font_size'])
                                except:
                                    font = ImageFont.load_default()
                                
                                # Draw text
                                draw.text(
                                    (element['actual_x'], element['actual_y']),
                                    text,
                                    fill=element['color'],
                                    font=font,
                                    anchor="mm"
                                )
                            
                            # Save the certificate
                            filename = f"certificate_{row_idx-1}.png"
                            
                            # Use name if available (assume first column is name)
                            if row[0]:
                                name = str(row[0]).replace(" ", "_")
                                filename = f"{name}_certificate.png"
                            
                            # Save to temporary file
                            cert_path = os.path.join(temp_dir, filename)
                            certificate.save(cert_path)
                            
                            # Add to zip
                            zip_file.write(cert_path, filename)
                            
                            # Store for email
                            if 'email_column' in st.session_state:
                                email_idx = st.session_state.excel_headers.index(st.session_state.email_column)
                                email = row[email_idx]
                                if email:
                                    cert_files[email] = cert_path
                            
                            certificate_count += 1
                        
                        st.session_state.certificate_files = cert_files
                    
                    # Prepare download link
                    zip_buffer.seek(0)
                    b64 = base64.b64encode(zip_buffer.read()).decode()
                    href = f'<a href="data:application/zip;base64,{b64}" download="certificates.zip" class="download-button">📥 Download All Certificates</a>'
                    
                    # Success message
                    st.success(f"✅ Successfully generated {certificate_count} certificates!")
                    st.markdown(href, unsafe_allow_html=True)
                    
                    # Set flag for email section
                    st.session_state.certificates_generated = True
                    
                except Exception as e:
                    st.error(f"Error generating certificates: {str(e)}")
        
        # Email section
        if 'email_column' in st.session_state and st.session_state.certificates_generated:
            st.header("Send Certificates by Email")
            
            if not st.session_state.sender_email or not st.session_state.email_password:
                st.warning("⚠️ Please enter your email credentials in the first tab to send emails")
            else:
                # Email settings
                subject = st.text_input("Email Subject", "Your Certificate")
                
                # Email body
                email_body = st.text_area(
                    "Email Body",
                    """Dear Participant,

Please find attached your certificate.

Best regards,
Certificate Team"""
                )
                
                # Send button
                if st.button("✉️ SEND ALL CERTIFICATES BY EMAIL", use_container_width=True):
                    st.info("Note: This functionality would need to be implemented on a secure server. For security reasons, the email sending is disabled in this demo but would work in a production environment.")
                    
                    # In a production environment, this would be the email sending code:
                    """
                    try:
                        # Set up server
                        context = ssl.create_default_context()
                        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                            server.login(st.session_state.sender_email, st.session_state.email_password)
                            
                            # Track progress
                            sent_count = 0
                            failed_emails = []
                            
                            # Send each email
                            for email, cert_path in st.session_state.certificate_files.items():
                                # Create message
                                msg = MIMEMultipart()
                                msg['From'] = st.session_state.sender_email
                                msg['To'] = email
                                msg['Subject'] = subject
                                
                                # Add body
                                msg.attach(MIMEText(email_body, 'plain'))
                                
                                # Attach certificate
                                with open(cert_path, 'rb') as f:
                                    attachment = MIMEApplication(f.read(), Name=os.path.basename(cert_path))
                                attachment['Content-Disposition'] = f'attachment; filename="{os.path.basename(cert_path)}"'
                                msg.attach(attachment)
                                
                                # Send email
                                server.sendmail(st.session_state.sender_email, email, msg.as_string())
                                sent_count += 1
                                
                            st.success(f"✅ Successfully sent {sent_count} emails!")
                    except Exception as e:
                        st.error(f"Error sending emails: {str(e)}")
                    """

# Add custom CSS for better styling
st.markdown("""
<style>
.download-button {
    display: inline-block;
    padding: 12px 24px;
    background-color: #4CAF50;
    color: white !important;
    text-decoration: none;
    font-weight: bold;
    border-radius: 4px;
    text-align: center;
    margin: 20px 0;
}
.download-button:hover {
    background-color: #45a049;
}
</style>
""", unsafe_allow_html=True)