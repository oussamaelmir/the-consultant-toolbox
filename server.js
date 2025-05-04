// server.js
const express = require('express');
const path = require('path');
const app = express();

// Serve all static files from wwwroot
app.use(express.static(path.join(__dirname, 'wwwroot')));

// Listen on Azure’s assigned port
const port = process.env.PORT || 8080;
app.listen(port, () =>
  console.log(`Static server running on port ${port}`)
);
