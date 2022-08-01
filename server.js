const express = require("express");
const app = express();
const server = require("http").createServer(app);
const io = require("socket.io")(server);
const fetch = require("node-fetch"); 

app.get("/", (req, res) => {
  res.sendFile(__dirname + "/soplaya_PianAtt/index.html");
});
app.get("/index.html", (req, res) => {
  res.sendFile(__dirname + "/soplaya_PianAtt/index.html");
});
app.get("/index.js", (req, res) => {
  res.sendFile(__dirname + "/soplaya_PianAtt/index.js");
});
app.get("/popup.html", (req, res) => {
  res.sendFile(__dirname + "/soplaya_PianAtt/popup.html");
});
app.get("/popup.js", (req, res) => {
  res.sendFile(__dirname + "/soplaya_PianAtt/popup.js");
});
app.get("/tableau.extensions.1.latest.js", (req, res) => {
  res.sendFile(__dirname + "/soplaya_PianAtt/tableau.extensions.1.latest.js");
});




app.get("/index_budget.html", (req, res) => {
  res.sendFile(__dirname + "/soplaya_PianAtt/index_budget.html");
});
app.get("/index_budget.js", (req, res) => {
  res.sendFile(__dirname + "/soplaya_PianAtt/index_budget.js");
});
app.get("/popup_budget.html", (req, res) => {
  res.sendFile(__dirname + "/soplaya_PianAtt/popup_budget.html");
});
app.get("/popup_budget.js", (req, res) => {
  res.sendFile(__dirname + "/soplaya_PianAtt/popup_budget.js");
});

//refer to https://tableau.github.io/extensions-api/docs/trex_oauth.html
//how to manage OAuth

app.get("/oathcomplete", async (req, res) => {
  const token=req.query.access_token;
  const socketid = req.query.state;
  console.log("token  " +token);
  console.log("socketid  " +socketid);
  if(token)
    if (socketid)
     io.to(socketid).emit("signedin", token);
  

  
  
  res.sendFile(__dirname + "/soplaya_PianAtt/complete.html");
  
});

app.get("/uncomplete", (req, res) => {
  res.sendFile(__dirname + "/soplaya_PianAtt/uncomplete.html");
});

server.listen(process.env.PORT, () => {
  console.log("Your app is listening on port " + server.address().port);
});
