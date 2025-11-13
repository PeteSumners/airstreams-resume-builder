# Resume Builder - JavaScript Version

A client-side resume builder that converts JSON resume data to Word documents. Works entirely in the browser with no backend required.

## Features

- 📄 Load resume data from JSON files
- 👁️ Live preview of resume content
- 📥 Export to Word (.docx) format
- 🚀 No server required - runs entirely in the browser
- 🌐 Can be hosted on GitHub Pages for free

## Local Usage

Simply open `index.html` in your web browser. That's it!

Or use a local server:

```bash
# Python 3
python -m http.server 8000

# Then visit http://localhost:8000
```

## JSON Format

Your JSON file should follow this structure:

```json
{
  "contact": {
    "name": "John Doe",
    "email": "john@example.com",
    "phone": "(555) 123-4567",
    "location": "City, State",
    "linkedin": "linkedin.com/in/johndoe",
    "github": "github.com/johndoe"
  },
  "objective": "Your career objective here...",
  "skills": ["Skill 1", "Skill 2", "Skill 3"],
  "certificates": ["Certificate 1", "Certificate 2"],
  "education": [
    {
      "institution": "University Name",
      "degree": "Bachelor of Science",
      "dates": "2020-2024",
      "location": "City, State",
      "details": ["Detail 1", "Detail 2"]
    }
  ],
  "experience": [
    {
      "company": "Company Name",
      "title": "Job Title",
      "dates": "Jan 2020 - Present",
      "responsibilities": [
        "Responsibility 1",
        "Responsibility 2"
      ]
    }
  ]
}
```

## Deploying to GitHub Pages

1. Create a new GitHub repository
2. Upload the contents of this folder to the repository
3. Go to repository Settings → Pages
4. Under "Source", select "main" branch
5. Click Save
6. Your resume builder will be live at: `https://yourusername.github.io/repository-name/`

### Quick Deploy Steps

```bash
# Initialize git (if not already done)
git init

# Add all files
git add .

# Commit
git commit -m "Initial commit of resume builder"

# Add your GitHub repository as remote
git remote add origin https://github.com/yourusername/repository-name.git

# Push to GitHub
git push -u origin main
```

Then enable GitHub Pages in your repository settings!

## Libraries Used

- [docx.js](https://docx.js.org/) - Word document generation
- [FileSaver.js](https://github.com/eligrey/FileSaver.js/) - File download functionality

Both libraries are loaded from CDN, so no installation required.

## Browser Compatibility

Works in all modern browsers:
- Chrome/Edge (recommended)
- Firefox
- Safari

## Troubleshooting

**Issue**: Export button doesn't work
- Check browser console for errors
- Ensure you're using a modern browser
- Try clearing cache and reloading

**Issue**: Preview doesn't show
- Verify your JSON file is valid (use a JSON validator)
- Check that all required fields are present
- Look for syntax errors in the JSON

## License

Free to use and modify as needed.
