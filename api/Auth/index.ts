import { AzureFunction, Context, HttpRequest } from "@azure/functions"

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
  //context.log('HTTP trigger function processed a request.');
  const name1 = req.query.name;

  // Roles -'Guest','User','Admin'
  const RoleMapping = [
    { user: "admin@sha2013.onmicrosoft.com", role: ["Admin"] },
    { user: "varsha@no2labs.onmicrosoft.com", role: ["User"] },
    { user: "tim@nudecision.com", role: ["Admin"] },
    { user: "tim@mtjohnston.net", role: ["User"] },
  ];

  const userroles = RoleMapping.filter(function (rl) {
    return rl.user.toLowerCase() == name1.toLowerCase();
  });

  const responseMessage = userroles.length > 0 ? userroles[0].role : []; // no authorized roles = []

  //const responseMessage="found1 "+name;
  context.res = {
    // status: 200, /* Defaults to 200 */
    // body: "found1 "+name1
    body: responseMessage,
  };
};
export default httpTrigger;
