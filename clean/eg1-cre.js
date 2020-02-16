const {
  google
} = require("googleapis");
const auth = require("./credentials-load");

async function run() {
  //create sheets client
  const sheets = google.sheets({
    version: "v4",
    auth
  });
  //get a range of values
  const res = await sheets.spreadsheets.create({
    resource: {
      // TODO: Add desired properties to the request body.
    },
  }, function (err, response) {
    if (err) {
      console.error(err);
      return;
    }

    console.log(response);
    // TODO: Change code below to process the `response` object:
    console.log(JSON.stringify(response, null, 2));
  });
}

run().catch(err => console.error("ERR", err));