const express = require('express');
const multer = require('multer');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, AlignmentType, LevelFormat } = require('docx');
const pdfParse = require('pdf-parse');
const mammoth = require('mammoth');

const app = express();
const PORT = 3000;

// Middleware
const allowedOrigins = [
    'https://eliteresumes.in',
    'https://www.eliteresumes.in',
    'http://localhost:3000',
    'http://localhost:5000',
    'http://localhost',
    'https://localhost'
];

app.use(cors({
    origin: function(origin, callback) {
        // Allow requests with no origin (mobile apps, curl requests)
        if (!origin) return callback(null, true);
        
        // Allow all origins for now - can be restricted later
        callback(null, true);
    },
    methods: ['GET', 'POST', 'OPTIONS', 'PUT', 'DELETE'],
    credentials: false,
    optionsSuccessStatus: 200,
    allowedHeaders: ['Content-Type', 'Authorization']
}));

app.use(express.json());
app.use(express.static('public'));

// Add preflight handling
app.options('*', cors());

// Configure multer for file uploads
const storage = multer.memoryStorage();
const upload = multer({ 
    storage: storage,
    limits: { fileSize: 10 * 1024 * 1024 } // 10MB limit
});

// Serve the HTML file
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Main conversion endpoint
app.post('/convert', upload.single('resume'), async (req, res) => {
    try {
        const { apiKey } = req.body;
        const file = req.file;

        if (!file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }

        if (!apiKey) {
            return res.status(400).json({ error: 'API key is required' });
        }

        console.log(`Processing file: ${file.originalname}`);

        // Step 1: Extract text from the uploaded file
        let resumeText = '';
        
        if (file.mimetype === 'application/pdf') {
            // Extract text from PDF
            const pdfData = await pdfParse(file.buffer);
            resumeText = pdfData.text;
        } else if (file.mimetype === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
                   file.mimetype === 'application/msword') {
            // Extract text from Word document
            const result = await mammoth.extractRawText({ buffer: file.buffer });
            resumeText = result.value;
        } else if (file.mimetype === 'text/plain') {
            // Plain text file
            resumeText = file.buffer.toString('utf-8');
        } else {
            return res.status(400).json({ error: 'Unsupported file type' });
        }

        console.log('Text extracted successfully');

        // Step 2: Call ChatGPT API to extract and structure resume data
        const resumeData = await callChatGPTAPI(apiKey, resumeText);
        console.log('Resume data structured successfully');

        // Step 3: Generate Word document
        const docBuffer = await generateWordDocument(resumeData);
        console.log('Word document generated successfully');

        // Send the document
        res.set({
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'Content-Disposition': `attachment; filename="${file.originalname.replace(/\.[^/.]+$/, '')}_converted.docx"`
        });
        res.send(docBuffer);

    } catch (error) {
        console.error('Conversion error:', error);
        res.status(500).json({ error: error.message });
    }
});

// Function to call ChatGPT API
async function callChatGPTAPI(apiKey, resumeText) {
    const fetch = (await import('node-fetch')).default;
    
    const prompt = `Convert the following resume to a specific template format. Extract ALL information and return it as a JSON object with this EXACT structure:

{
    "name": "FULL NAME IN CAPS",
    "location": "City, Country",
    "phone": "+XX XXXXXXXXXX",
    "email": "email@example.com",
    "summary": ["paragraph1", "paragraph2"],
    "experience": [
        {
            "title": "Job Title",
            "dates": "MMM YYYY â€“ Present/MMM YYYY",
            "company": "Company Name, Location",
            "responsibilities": ["responsibility1", "responsibility2", "..."]
        }
    ],
    "education": [
        {
            "degree": "Degree Name",
            "institution": "Institution Name, Location",
            "year": "YYYY or Pursuing"
        }
    ],
    "certifications": ["cert1", "cert2"],
    "skills": {
        "technical": "comma-separated skills",
        "core": "comma-separated competencies"
    },
    "achievements": ["achievement1", "achievement2"],
    "personal": {
        "nationality": "Country",
        "languages": "Language1 (Level), Language2 (Level)",
        "visaStatus": "Status if mentioned",
        "other": ["other detail 1", "other detail 2"]
    }
}

IMPORTANT RULES:
- Extract ALL information from the resume
- Convert name to ALL CAPS
- Keep professional summary in exactly 2 paragraphs (combine if more, split if one)
- Include ALL job experiences with ALL bullet points
- Preserve ALL dates and details exactly as mentioned
- If a section doesn't exist, use empty array or empty string
- Return ONLY the JSON object, no other text or markdown

Resume to convert:
${resumeText}`;

    try {
        console.log('Calling ChatGPT API...');
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify({
                model: 'gpt-4o-mini',
                messages: [{
                    role: 'user',
                    content: prompt
                }],
                temperature: 0.3,
                max_tokens: 4000
            })
        });

        if (!response.ok) {
            const errorData = await response.json();
            const errorMessage = errorData.error?.message || `HTTP ${response.status}`;
            throw new Error(`ChatGPT API error: ${errorMessage}`);
        }

        const data = await response.json();
        const responseText = data.choices[0].message.content;
        
        // Parse JSON from ChatGPT's response
        try {
            return JSON.parse(responseText);
        } catch (e) {
            // Try to extract JSON if wrapped in markdown
            const jsonMatch = responseText.match(/\{[\s\S]*\}/);
            if (jsonMatch) {
                return JSON.parse(jsonMatch[0]);
            }
            throw new Error('Failed to parse ChatGPT response as JSON');
        }

    } catch (error) {
        console.error('ChatGPT API error:', error);
        throw error;
    }
}

// Function to generate Word document - EXACT COPY FROM ORIGINAL
async function generateWordDocument(data) {
    const doc = new Document({
        numbering: {
            config: [{
                reference: "bullet-list",
                levels: [{
                    level: 0,
                    format: LevelFormat.BULLET,
                    text: "â€¢",
                    alignment: AlignmentType.LEFT,
                    style: {
                        paragraph: {
                            indent: { left: 720, hanging: 360 }
                        }
                    }
                }]
            }]
        },
        sections: [{
            properties: {
                page: {
                    margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
                }
            },
            children: [
                // Name
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 100 },
                    children: [
                        new TextRun({
                            text: data.name || "NAME",
                            size: 32,
                            bold: true
                        })
                    ]
                }),
                
                // Location and phone
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 50 },
                    children: [
                        new TextRun({
                            text: `${data.location || "Location"}|${data.phone || "Phone"}`,
                            size: 22
                        })
                    ]
                }),
                
                // Email
                new Paragraph({
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 200 },
                    children: [
                        new TextRun({
                            text: data.email || "email@example.com",
                            size: 22,
                            underline: {}
                        })
                    ]
                }),
                
                // Professional Summary
                ...(data.summary || []).map((para, index) => 
                    new Paragraph({
                        alignment: AlignmentType.JUSTIFIED,
                        spacing: { after: index === data.summary.length - 1 ? 240 : 120 },
                        children: [
                            new TextRun({
                                text: para,
                                size: 22
                            })
                        ]
                    })
                ),
                
                // EXPERIENCE Section
                new Paragraph({
                    spacing: { before: 120, after: 120 },
                    children: [
                        new TextRun({
                            text: "EXPERIENCE",
                            size: 24,
                            bold: true
                        })
                    ]
                }),
                
                // Jobs
                ...(data.experience || []).flatMap((job, jobIndex) => [
                    new Paragraph({
                        spacing: { after: 60 },
                        children: [
                            new TextRun({ 
                                text: job.title || "Job Title", 
                                bold: true, 
                                size: 22 
                            }),
                            new TextRun({ 
                                text: `                                         ${job.dates || "Dates"}`, 
                                size: 22 
                            })
                        ]
                    }),
                    new Paragraph({
                        spacing: { after: 80 },
                        children: [
                            new TextRun({ 
                                text: job.company || "Company Name", 
                                size: 22 
                            })
                        ]
                    }),
                    ...(job.responsibilities || []).map((resp, respIndex) => 
                        new Paragraph({
                            numbering: { reference: "bullet-list", level: 0 },
                            spacing: { 
                                after: respIndex === job.responsibilities.length - 1 && 
                                       jobIndex < data.experience.length - 1 ? 120 : 60 
                            },
                            children: [
                                new TextRun({ 
                                    text: resp, 
                                    size: 22 
                                })
                            ]
                        })
                    )
                ]),
                
                // EDUCATION Section
                new Paragraph({
                    spacing: { before: 120, after: 120 },
                    children: [
                        new TextRun({
                            text: "EDUCATON",
                            size: 24,
                            bold: true
                        })
                    ]
                }),
                
                // Education entries
                ...(data.education || []).flatMap((edu, index) => [
                    new Paragraph({
                        spacing: { after: 60 },
                        children: [
                            new TextRun({ 
                                text: edu.degree || "Degree", 
                                bold: true, 
                                size: 22 
                            })
                        ]
                    }),
                    new Paragraph({
                        spacing: { after: index === data.education.length - 1 ? 180 : 120 },
                        children: [
                            new TextRun({ 
                                text: `${edu.institution || "Institution"} | ${edu.year || "Year"}`, 
                                size: 22 
                            })
                        ]
                    })
                ]),
                
                // CERTIFICATIONS Section (if exists)
                ...(data.certifications && data.certifications.length > 0 ? [
                    new Paragraph({
                        spacing: { before: 120, after: 120 },
                        children: [
                            new TextRun({
                                text: "CERTIFICATIONS",
                                size: 24,
                                bold: true
                            })
                        ]
                    }),
                    ...data.certifications.map((cert, index) => 
                        new Paragraph({
                            numbering: { reference: "bullet-list", level: 0 },
                            spacing: { after: index === data.certifications.length - 1 ? 180 : 60 },
                            children: [
                                new TextRun({ 
                                    text: cert, 
                                    size: 22 
                                })
                            ]
                        })
                    )
                ] : []),
                
                // KEY ACHIEVEMENTS Section (if exists)
                ...(data.achievements && data.achievements.length > 0 ? [
                    new Paragraph({
                        spacing: { before: 120, after: 120 },
                        children: [
                            new TextRun({
                                text: "KEY ACHIEVEMENTS",
                                size: 24,
                                bold: true
                            })
                        ]
                    }),
                    ...data.achievements.map((achievement, index) => 
                        new Paragraph({
                            numbering: { reference: "bullet-list", level: 0 },
                            spacing: { after: index === data.achievements.length - 1 ? 180 : 60 },
                            children: [
                                new TextRun({ 
                                    text: achievement, 
                                    size: 22 
                                })
                            ]
                        })
                    )
                ] : []),
                
                // SKILLS Section
                new Paragraph({
                    spacing: { before: 120, after: 120 },
                    children: [
                        new TextRun({
                            text: "SKILLS",
                            size: 24,
                            bold: true
                        })
                    ]
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { after: 100 },
                    children: [
                        new TextRun({ 
                            text: "Technical skills: ", 
                            bold: true, 
                            size: 22 
                        }),
                        new TextRun({ 
                            text: data.skills?.technical || "Skills to be added", 
                            size: 22 
                        })
                    ]
                }),
                new Paragraph({
                    alignment: AlignmentType.JUSTIFIED,
                    spacing: { after: 180 },
                    children: [
                        new TextRun({ 
                            text: "Core competencies: ", 
                            bold: true, 
                            size: 22 
                        }),
                        new TextRun({ 
                            text: data.skills?.core || "Competencies to be added", 
                            size: 22 
                        })
                    ]
                }),
                
                // PERSONAL DETAILS Section
                new Paragraph({
                    spacing: { before: 120, after: 120 },
                    children: [
                        new TextRun({
                            text: "PERSONAL DETAILS",
                            size: 24,
                            bold: true
                        })
                    ]
                }),
                new Paragraph({
                    numbering: { reference: "bullet-list", level: 0 },
                    spacing: { after: 60 },
                    children: [
                        new TextRun({ 
                            text: `Nationality: ${data.personal?.nationality || "To be added"}`, 
                            size: 22 
                        })
                    ]
                }),
                ...(data.personal?.languages ? [
                    new Paragraph({
                        numbering: { reference: "bullet-list", level: 0 },
                        spacing: { after: 60 },
                        children: [
                            new TextRun({ 
                                text: `Languages: ${data.personal.languages}`, 
                                size: 22 
                            })
                        ]
                    })
                ] : []),
                ...(data.personal?.visaStatus ? [
                    new Paragraph({
                        numbering: { reference: "bullet-list", level: 0 },
                        spacing: { after: 60 },
                        children: [
                            new TextRun({ 
                                text: `Visa Status: ${data.personal.visaStatus}`, 
                                size: 22 
                            })
                        ]
                    })
                ] : []),
                ...(data.personal?.other || []).map((detail, index) => 
                    new Paragraph({
                        numbering: { reference: "bullet-list", level: 0 },
                        spacing: { after: 60 },
                        children: [
                            new TextRun({ 
                                text: detail, 
                                size: 22 
                            })
                        ]
                    })
                )
            ]
        }]
    });

    return await Packer.toBuffer(doc);
}

// Start server
app.listen(PORT, () => {
    console.log(`Resume converter server running at http://localhost:${PORT}`);
    console.log('\nInstructions:');
    console.log('1. Make sure you have installed required packages:');
    console.log('   npm install express multer cors docx pdf-parse mammoth node-fetch');
    console.log('2. Place the HTML file in a "public" folder');
    console.log('3. Access the application at http://localhost:3000');
});