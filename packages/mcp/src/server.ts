import { FastMCP } from "fastmcp";
import { registerAuthTools } from "./tools/auth.js";
import { registerMailTools } from "./tools/mail.js";

const server = new FastMCP({
  name: "outlook-toolkit",
  version: "0.1.0",
});

registerAuthTools(server);
registerMailTools(server);

export default server;
