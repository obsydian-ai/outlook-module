const httpServer = require('http-server');
const port = process.env.PORT || 3000;

const server = httpServer.createServer({
  root: '.',
  cors: true,
  cache: -1
});

server.listen(port, () => {
  console.log(`Server running on port ${port}`);
});

