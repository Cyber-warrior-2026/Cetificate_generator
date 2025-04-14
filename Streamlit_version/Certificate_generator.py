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
import json
import time
import re

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
if 'template_file' not in st.session_state:
    st.session_state.template_file = None
if 'excel_headers' not in st.session_state:
    st.session_state.excel_headers = []
if 'certificates_generated' not in st.session_state:
    st.session_state.certificates_generated = False
if 'certificate_files' not in st.session_state:
    st.session_state.certificate_files = {}
if 'active_tab' not in st.session_state:
    st.session_state.active_tab = 0
if 'errors' not in st.session_state:
    st.session_state.errors = []
if 'last_drag_update' not in st.session_state:
    st.session_state.last_drag_update = time.time()

# Function to validate email
def is_valid_email(email):
    pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    return re.match(pattern, email) is not None

# Function to test email connection
def test_email_connection(email, password):
    try:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(email, password)
        return True, "Connection successful"
    except smtplib.SMTPAuthenticationError:
        return False, "Authentication failed. Check your email and app password."
    except smtplib.SMTPException as e:
        return False, f"SMTP error: {str(e)}"
    except Exception as e:
        return False, f"Connection error: {str(e)}"

# Title and description
st.title("Certificate Generator")
st.markdown("""
Generate beautiful certificates from an Excel list of participants.
Upload your certificate template and Excel data, position your text fields using drag-and-drop, and download all certificates in one click!
""")

# Show errors if any
if st.session_state.errors:
    with st.expander("Error Log", expanded=True):
        for error in st.session_state.errors:
            st.error(error)
        if st.button("Clear Errors"):
            st.session_state.errors = []
            st.experimental_rerun()

# Create tabs for workflow
tab_names = ["1. Upload Files", "2. Design Certificate", "3. Generate & Send"]
tabs = st.tabs(tab_names)

# Tab 1: Upload Files
with tabs[0]:
    st.header("Upload Your Files")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Upload certificate template
        st.subheader("Certificate Template")
        template_file = st.file_uploader("Upload certificate template image", type=["png", "jpg", "jpeg"])
        
        if template_file:
            try:
                # Display the template
                image = Image.open(template_file)
                st.session_state.template_size = image.size
                st.session_state.template_file = template_file
                st.image(image, caption=f"Certificate Template ({image.width}x{image.height} px)", use_column_width=True)
                st.success(f"✅ Template loaded: {image.width}x{image.height} pixels")
            except Exception as e:
                error_msg = f"Error loading template image: {str(e)}"
                st.session_state.errors.append(error_msg)
                st.error(error_msg)
    
    with col2:
        # Upload Excel file
        st.subheader("Participant Data")
        excel_file = st.file_uploader("Upload Excel file with participant data", type=["xlsx", "xls"])
        
        if excel_file:
            try:
                # Load Excel data
                wb = openpyxl.load_workbook(excel_file)
                sheet = wb.active
                
                if sheet.max_row < 2:
                    st.warning("Excel file appears to be empty. Please check your data.")
                else:
                    headers = [str(cell.value) for cell in sheet[1] if cell.value]
                    
                    if not headers:
                        st.warning("No headers found in Excel file. Please check your data.")
                    else:
                        st.session_state.excel_headers = headers
                        
                        # Show Excel preview
                        preview_data = []
                        for i, row in enumerate(sheet.iter_rows(values_only=True)):
                            if i == 0:  # Headers
                                preview_data.append(row[:len(headers)])
                            elif i < 6:  # First 5 rows
                                preview_data.append(row[:len(headers)])
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
                                sender_email = st.text_input("Your email address", 
                                                      value=st.session_state.get('sender_email', ''))
                                if sender_email and not is_valid_email(sender_email):
                                    st.warning("Please enter a valid email address")
                                else:
                                    st.session_state.sender_email = sender_email
                            
                            with col_b:
                                st.session_state.email_password = st.text_input(
                                    "App password", 
                                    value=st.session_state.get('email_password', ''),
                                    type="password",
                                    help="To send emails using Gmail, you'll need an App Password. This is a 16-character code that gives permission to apps."
                                )
                            
                            # Test connection button
                            if st.session_state.get('sender_email') and st.session_state.get('email_password'):
                                if st.button("Test Email Connection"):
                                    success, message = test_email_connection(
                                        st.session_state.sender_email,
                                        st.session_state.email_password
                                    )
                                    if success:
                                        st.success(message)
                                    else:
                                        st.error(message)
                                        st.info("If using Gmail, ensure you've set up an App Password: [Learn how](https://support.google.com/accounts/answer/185833)")
                        
            except Exception as e:
                error_msg = f"Error loading Excel file: {str(e)}"
                st.session_state.errors.append(error_msg)
                st.error(error_msg)
    
    # Navigation buttons
    if st.session_state.get('template_size') and st.session_state.excel_headers:
        if st.button("Continue to Design ➡️", use_container_width=True):
            st.session_state.active_tab = 1
            st.experimental_rerun()

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
            
            # Add field button
            if st.button("Add Field", use_container_width=True):
                # Default position at center
                x_pos = 50
                y_pos = 50
                
                # Calculate actual coordinates
                actual_x = int(st.session_state.template_size[0] * x_pos / 100)
                actual_y = int(st.session_state.template_size[1] * y_pos / 100)
                
                # Add to session state with width and height for dragging
                st.session_state.text_elements.append({
                    'field': field,
                    'font_size': font_size,
                    'color': color,
                    'x_pos': x_pos,
                    'y_pos': y_pos,
                    'actual_x': actual_x,
                    'actual_y': actual_y,
                    'width': font_size * len(field) * 0.6,  # Approximate width
                    'height': font_size * 1.2,  # Approximate height
                    'id': len(st.session_state.text_elements)
                })
                st.success(f"✅ Added field: {field}")
                st.experimental_rerun()
            
            # List current fields
            st.subheader("Current Fields")
            if not st.session_state.text_elements:
                st.info("No fields added yet")
            else:
                for i, element in enumerate(st.session_state.text_elements):
                    col_a, col_b, col_c = st.columns([2, 1, 1])
                    with col_a:
                        st.write(f"**{element['field']}** (Size: {element['font_size']})")
                    with col_b:
                        if st.button("Edit", key=f"edit_{i}"):
                            st.session_state.editing_element = i
                            st.experimental_rerun()
                    with col_c:
                        if st.button("🗑️", key=f"del_{i}"):
                            st.session_state.text_elements.pop(i)
                            st.experimental_rerun()
                    
                    # Show edit form if editing this element
                    if st.session_state.get('editing_element') == i:
                        with st.form(f"edit_element_{i}"):
                            element['font_size'] = st.slider("Font size", 
                                                           min_value=10, 
                                                           max_value=100,
                                                           value=element['font_size'],
                                                           key=f"edit_size_{i}")
                            
                            element['color'] = st.color_picker("Text color", 
                                                             element['color'],
                                                             key=f"edit_color_{i}")
                            
                            element['x_pos'] = st.slider("X position (%)", 
                                                       min_value=0, 
                                                       max_value=100,
                                                       value=element['x_pos'],
                                                       key=f"edit_x_{i}")
                            
                            element['y_pos'] = st.slider("Y position (%)", 
                                                       min_value=0, 
                                                       max_value=100,
                                                       value=element['y_pos'],
                                                       key=f"edit_y_{i}")
                            
                            # Update actual coordinates
                            element['actual_x'] = int(st.session_state.template_size[0] * element['x_pos'] / 100)
                            element['actual_y'] = int(st.session_state.template_size[1] * element['y_pos'] / 100)
                            
                            col1, col2 = st.columns(2)
                            with col1:
                                if st.form_submit_button("Save Changes"):
                                    del st.session_state.editing_element
                                    st.experimental_rerun()
                            with col2:
                                if st.form_submit_button("Cancel"):
                                    del st.session_state.editing_element
                                    st.experimental_rerun()
            
            # Clear all button
            if st.session_state.text_elements:
                if st.button("Clear All Fields", use_container_width=True):
                    st.session_state.text_elements = []
                    st.experimental_rerun()
        
        with col2:
            st.subheader("Certificate Preview")
            
            # Load the template
            if st.session_state.template_file:
                try:
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
                    
                    # Convert to base64 for the canvas
                    img_buffer = io.BytesIO()
                    image.save(img_buffer, format="PNG")
                    img_data = base64.b64encode(img_buffer.getvalue()).decode("utf-8")
                    
                    # Show preview with interactive canvas
                    st.markdown("### Drag fields to position them on the certificate")
                    
                    # Create interactive canvas with HTML/JS
                    canvas_html = f"""
                    <div style="position: relative;">
                        <img src="data:image/png;base64,{img_data}" style="width: 100%; max-width: 800px;" id="certificate-img">
                        <div id="canvas-container" style="position: absolute; top: 0; left: 0; width: 100%; height: 100%;">
                            <!-- Elements will be dynamically added here -->
                        </div>
                    </div>
                    
                    <script>
                    // Wait for the image to load
                    document.getElementById('certificate-img').onload = function() {{
                        const container = document.getElementById('canvas-container');
                        container.style.width = this.offsetWidth + 'px';
                        container.style.height = this.offsetHeight + 'px';
                        
                        // Get current elements from JSON
                        const elements = {json.dumps(st.session_state.text_elements)};
                        const imageWidth = {st.session_state.template_size[0]};
                        const imageHeight = {st.session_state.template_size[1]};
                        const displayWidth = this.offsetWidth;
                        const displayHeight = this.offsetHeight;
                        
                        // Scale factors
                        const scaleX = displayWidth / imageWidth;
                        const scaleY = displayHeight / imageHeight;
                        
                        // Add each element
                        elements.forEach((el, index) => {{
                            // Create draggable div
                            const elem = document.createElement('div');
                            elem.id = 'element-' + el.id;
                            elem.className = 'draggable-element';
                            elem.style.position = 'absolute';
                            elem.style.cursor = 'move';
                            elem.style.backgroundColor = 'rgba(255,255,255,0.3)';
                            elem.style.border = '1px dashed #333';
                            elem.style.padding = '5px';
                            elem.style.borderRadius = '4px';
                            elem.style.fontSize = Math.max(10, el.font_size * scaleY * 0.8) + 'px';
                            elem.style.color = el.color;
                            elem.style.left = (el.actual_x * scaleX - (el.width * scaleX / 2)) + 'px';
                            elem.style.top = (el.actual_y * scaleY - (el.height * scaleY / 2)) + 'px';
                            elem.style.minWidth = '50px';
                            elem.style.textAlign = 'center';
                            elem.innerHTML = el.field;
                            
                            // Add element to container
                            container.appendChild(elem);
                            
                            // Make draggable
                            let isDragging = false;
                            let offsetX, offsetY;
                            
                            elem.addEventListener('mousedown', function(e) {{
                                isDragging = true;
                                offsetX = e.clientX - this.getBoundingClientRect().left;
                                offsetY = e.clientY - this.getBoundingClientRect().top;
                                this.style.zIndex = 1000;
                            }});
                            
                            document.addEventListener('mousemove', function(e) {{
                                if (!isDragging) return;
                                
                                const x = e.clientX - offsetX - container.getBoundingClientRect().left;
                                const y = e.clientY - offsetY - container.getBoundingClientRect().top;
                                
                                // Keep within bounds
                                const boundedX = Math.max(0, Math.min(container.offsetWidth - 50, x));
                                const boundedY = Math.max(0, Math.min(container.offsetHeight - 20, y));
                                
                                elem.style.left = boundedX + 'px';
                                elem.style.top = boundedY + 'px';
                                
                                // Calculate position percentage
                                const centerX = boundedX + (parseInt(elem.style.width) || 50) / 2;
                                const centerY = boundedY + (parseInt(elem.style.height) || 20) / 2;
                                
                                // Update hidden form field
                                const xPos = Math.round((centerX / container.offsetWidth) * 100);
                                const yPos = Math.round((centerY / container.offsetHeight) * 100);
                                
                                // Send data to Streamlit using component communication
                                if (window.Streamlit) {{
                                    const data = {{
                                        elementId: el.id,
                                        xPos: xPos,
                                        yPos: yPos
                                    }};
                                    window.Streamlit.setComponentValue(JSON.stringify(data));
                                }}
                            }});
                            
                            document.addEventListener('mouseup', function() {{
                                if (isDragging) {{
                                    isDragging = false;
                                    elem.style.zIndex = 'auto';
                                }}
                            }});
                        }});
                    }};
                    </script>
                    """
                    
                    # Use a component to handle the drag events
                    component_value = st.components.v1.html(canvas_html, height=600)
                    
                    # Process drag updates
                    if component_value:
                        try:
                            data = json.loads(component_value)
                            # Update position in session state
                            for i, element in enumerate(st.session_state.text_elements):
                                if element['id'] == data['elementId']:
                                    # Update percentages and actual coordinates
                                    element['x_pos'] = data['xPos']
                                    element['y_pos'] = data['yPos']
                                    element['actual_x'] = int(st.session_state.template_size[0] * data['xPos'] / 100)
                                    element['actual_y'] = int(st.session_state.template_size[1] * data['yPos'] / 100)
                                    break
                            
                            # Prevent too many reruns by limiting the update frequency
                            current_time = time.time()
                            if current_time - st.session_state.last_drag_update > 1.0:
                                st.session_state.last_drag_update = current_time
                                st.experimental_rerun()
                                
                        except json.JSONDecodeError:
                            pass
                        
                except Exception as e:
                    error_msg = f"Error generating preview: {str(e)}"
                    st.session_state.errors.append(error_msg)
                    st.error(error_msg)
            else:
                st.info("Please upload a template image in the previous tab")
        
        # Navigation
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("⬅️ Back to Upload", use_container_width=True):
                st.session_state.active_tab = 0
                st.experimental_rerun()
        with col_b:
            if len(st.session_state.text_elements) > 0:
                if st.button("Continue to Generate ➡️", use_container_width=True):
                    st.session_state.active_tab = 2
                    st.experimental_rerun()
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
                    excel_file = st.session_state.get('excel_file', None)
                    
                    if excel_file is None:
                        st.error("Excel file not found. Please go back to the first tab and upload it again.")
                    else:
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
                            
                            # Progress bar
                            progress_bar = st.progress(0)
                            total_rows = sheet.max_row - 1
                            
                            for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
                                # Update progress
                                progress_percent = min(int((row_idx-1) / total_rows * 100), 100)
                                progress_bar.progress(progress_percent)
                                
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
                                        try:
                                            # Try system fonts as fallback
                                            system_fonts = [
                                                "DejaVuSans.ttf", 
                                                "LiberationSans-Regular.ttf",
                                                "NotoSans-Regular.ttf",
                                                "Tahoma.ttf"
                                            ]
                                            
                                            for font_name in system_fonts:
                                                try:
                                                    font = ImageFont.truetype(font_name, element['font_size'])
                                                    break
                                                except:
                                                    continue
                                            
                                            if 'font' not in locals():
                                                font = ImageFont.load_default()
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
                                    if email_idx < len(row) and row[email_idx]:
                                        email = row[email_idx]
                                        if email and is_valid_email(email):
                                            cert_files[email] = cert_path
                                
                                certificate_count += 1
                            
                            st.session_state.certificate_files = cert_files
                        
                        # Complete the progress bar
                        progress_bar.progress(100)
                        
                        # Prepare download link
                        zip_buffer.seek(0)
                        b64 = base64.b64encode(zip_buffer.read()).decode()
                        href = f'<a href="data:application/zip;base64,{b64}" download="certificates.zip" class="download-button">📥 Download All Certificates</a>'
                        
                        # Success message
                        st.success(f"✅ Successfully generated {certificate_count} certificates!")
                        st.markdown(href, unsafe_allow_html=True)
                        
                        # Set flag for email section
                        st.session_state.certificates_generated = True
                
                except FileNotFoundError as e:
                    error_msg = "File not found. Please try uploading your files again."
                    st.session_state.errors.append(error_msg)
                    st.error(error_msg)
                except PermissionError as e:
                    error_msg = "Permission error. Cannot write to temporary directory."
                    st.session_state.errors.append(error_msg)
                    st.error(error_msg)
                except Exception as e:
                    error_msg = f"Error generating certificates: {str(e)}"
                    st.session_state.errors.append(error_msg)
                    st.error(error_msg)
        
        # Email section
        if 'email_column' in st.session_state and st.session_state.certificates_generated:
            st.header("Send Certificates by Email")
            
            if not st.session_state.get('sender_email') or not st.session_state.get('email_password'):
                st.warning("⚠️ Please enter your email credentials in the first tab to send emails")
            else:
                # Email settings
                with st.form("email_form"):
                    subject = st.text_input("Email Subject", "Your Certificate")
                    
                    # Email body
                    email_body = st.text_area(
                        "Email Body",
                        """Dear Participant,

Please find attached your certificate.

Best regards,
Certificate Team"""
                    )
                    
                    # Test email option
                    test_email = st.text_input("Send test email to (optional)", help="Enter an email address to send a test certificate")
                    
                    send_all = st.form_submit_button("✉️ SEND ALL CERTIFICATES BY EMAIL", use_container_width=True)
                    send_test = st.form_submit_button("🧪 SEND TEST EMAIL", use_container_width=True)
                
                if send_test and test_email:
                    if not is_valid_email(test_email):
                        st.error("Please enter a valid email address for the test")
                    else:
                        with st.spinner("Sending test email..."):
                            try:
                                # Set up server
                                context = ssl.create_default_context()
                                with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                                    server.login(st.session_state.sender_email, st.session_state.email_password)
                                    
                                    # Choose first certificate for test
                                    cert_path = next(iter(st.session_state.certificate_files.values()))
                                    
                                    # Create message
                                    msg = MIMEMultipart()
                                    msg['From'] = st.session_state.sender_email
                                    msg['To'] = test_email
                                    msg['Subject'] = subject + " (TEST)"
                                    
                                    # Add body
                                    msg.attach(MIMEText(email_body, 'plain'))
                                    
                                    # Attach certificate
                                    with open(cert_path, 'rb') as f:
                                        attachment = MIMEApplication(f.read(), Name=os.path.basename(cert_path))
                                    attachment['Content-Disposition'] = f'attachment; filename="{os.path.basename(cert_path)}"'
                                    msg.attach(attachment)
                                    
                                    # Send email
                                    # server.sendmail(st.session_state.sender_email, test_email, msg.as_string())
                                    # Send email
                                    server.sendmail(st.session_state.sender_email, test_email, msg.as_string())
                                    
                                    st.success(f"✅ Test email sent to {test_email}!")
                            except smtplib.SMTPAuthenticationError:
                                st.error("Authentication failed. Please check your email and app password.")
                                st.info("If using Gmail, ensure you've set up an App Password: [Learn how](https://support.google.com/accounts/answer/185833)")
                            except smtplib.SMTPRecipientsRefused:
                                st.error(f"Email {test_email} was refused by the server. Please check the address.")
                            except smtplib.SMTPException as e:
                                st.error(f"SMTP error: {str(e)}")
                            except FileNotFoundError:
                                st.error("Certificate file not found. Please generate certificates again.")
                            except Exception as e:
                                st.error(f"Error sending test email: {str(e)}")
                
                if send_all:
                    if not st.session_state.certificate_files:
                        st.error("No certificates with valid email addresses were found.")
                    else:
                        with st.spinner("Sending emails..."):
                            try:
                                # Set up server
                                context = ssl.create_default_context()
                                with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                                    server.login(st.session_state.sender_email, st.session_state.email_password)
                                    
                                    # Track progress
                                    sent_count = 0
                                    failed_emails = []
                                    
                                    # Progress bar
                                    progress_bar = st.progress(0)
                                    total_emails = len(st.session_state.certificate_files)
                                    
                                    # Send each email
                                    for i, (email, cert_path) in enumerate(st.session_state.certificate_files.items()):
                                        # Update progress
                                        progress_percent = min(int(i / total_emails * 100), 100)
                                        progress_bar.progress(progress_percent)
                                        
                                        try:
                                            # Check if file exists
                                            if not os.path.exists(cert_path):
                                                failed_emails.append(f"{email} (certificate file missing)")
                                                continue
                                                
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
                                            
                                        except Exception as e:
                                            failed_emails.append(f"{email} ({str(e)})")
                                    
                                    # Complete the progress bar
                                    progress_bar.progress(100)
                                    
                                    # Report results
                                    if failed_emails:
                                        st.warning(f"✅ Successfully sent {sent_count} emails, with {len(failed_emails)} failures.")
                                        with st.expander("Show failed emails"):
                                            for email in failed_emails:
                                                st.write(f"- {email}")
                                    else:
                                        st.success(f"✅ Successfully sent all {sent_count} emails!")
                            
                            except smtplib.SMTPAuthenticationError:
                                st.error("Authentication failed. Please check your email and app password.")
                                st.info("If using Gmail, ensure you've set up an App Password: [Learn how](https://support.google.com/accounts/answer/185833)")
                            except Exception as e:
                                st.error(f"Error sending emails: {str(e)}")

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
.draggable-element {
    transition: box-shadow 0.2s ease;
}
.draggable-element:hover {
    box-shadow: 0 0 10px rgba(0,0,0,0.3);
}
.stApp {
    max-width: 1200px;
    margin: 0 auto;
}
</style>
""", unsafe_allow_html=True)

# Fix for email upload persistence
if 'excel_file' in st.session_state and excel_file is not None:
    st.session_state.excel_file = excel_file