var graph = require('@microsoft/microsoft-graph-client');

module.exports = {
  getUserDetails: async function(accessToken) {
    const client = getAuthenticatedClient(accessToken);

    const user = await client.api('/me').get();
    return user;
  },

  getEvents: async function(accessToken) {
    const client = getAuthenticatedClient(accessToken);

//get current date time
//hardcoding dates right now. This needs to be corrected.
var dateTime = require('node-datetime');
var dt = dateTime.create();
var formatted = dt.format('Y-m-d');
var datestart = new Date();
datestart.setHours(5,30,0,0);
var  start = datestart.toISOString();

dateend = new Date();
var numberOfDaysToAdd = 5;
dateend.setDate(dateend.getDate() + numberOfDaysToAdd); 
dateend.setHours(5,29,59,59);
var end = dateend.toISOString();


console.log(formatted);
    const events = await client
    .api("/me/calendarview")
    .query({
//      startdatetime: "2019-07-23T00:00:00.0000000",
//      enddatetime: "2019-07-26T23:59:59.0000000"
      startdatetime: start,
      enddatetime: end
    }).select('subject,organizer,start,end')
      .orderby('end/dateTime ASC')
      .top(50)
      .header('Prefer', 'outlook.timezone="India Standard Time"')
      .get();
  console.log(events);
    return events;
  }
};

function getAuthenticatedClient(accessToken) {
  // Initialize Graph client
  const client = graph.Client.init({
    // Use the provided access token to authenticate
    // requests
    authProvider: (done) => {
      done(null, accessToken);
    }
  });

  return client;
}