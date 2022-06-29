import path from "path";
import * as restify from "restify";
import { AdaptiveCardsTabCommandHandler } from "./adaptiveCardsTabCommandHandler";
import { commandBot } from "./internal/initialize";

// This template uses `restify` to serve HTTP responses.
// Create a restify server.
const server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nBot Started, ${server.name} listening to ${server.url}`);
});

server.get('/images/*',
     restify.plugins.serveStaticFiles(path.resolve(__dirname, './images'))
);

// Register an API endpoint with `restify`. Teams sends messages to your application
// through this endpoint.
//
// The Teams Toolkit bot registration configures the bot with `/api/messages` as the
// Bot Framework endpoint. If you customize this route, update the Bot registration
// in `/templates/provision/bot.bicep`.
const acTabCommandHandler = new AdaptiveCardsTabCommandHandler();
server.post("/api/messages", async (req, res) => {
  await commandBot.requestHandler(req, res, async (context) => {
    await acTabCommandHandler.run(context);
  });
});
