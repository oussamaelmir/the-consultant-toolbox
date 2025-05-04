// server.js
const express = require('express');
const path = require('path');
const app = express();

// Serve everything directly from the root of the deployed folder
app.use(express.static(__dirname, {
    index: ['index.html'],      // still serve index.html for directories
    extensions: ['html'],       // try appending .html for extensionless requests
  }));
// If you ever want a catch-all (e.g. for SPA), you can uncomment this:
// app.get('*', (req, res) => res.sendFile(path.join(__dirname, 'index.html')));

const port = process.env.PORT || 8080;
app.listen(port, () => console.log(`Listening on ${port}`));
