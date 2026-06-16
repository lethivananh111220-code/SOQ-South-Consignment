import codecs

try:
    with open('app.js', 'rb') as f:
        raw_data = f.read()

    # Try decode with utf-8 first
    try:
        text = raw_data.decode('utf-8')
    except UnicodeDecodeError:
        text = raw_data.decode('utf-16', errors='ignore')

    # Save to utf-8 cleanly
    with codecs.open('app_utf8.js', 'w', encoding='utf-8') as f:
        f.write(text)
        
    print("Converted to app_utf8.js successfully.")
except Exception as e:
    print(f"Error: {e}")
