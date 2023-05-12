import { AzureFunction, Context, HttpRequest } from "@azure/functions"

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
  //context.log('HTTP trigger function processed a request.');
  const name1 = req.query.name;

  // //Roles -'Reader','Editor','Administrator'
  const RoleMapping = [
    { user: "admin@sha2013.onmicrosoft.com", role: ["Administrator"] },
    { user: "varsha@no2labs.onmicrosoft.com", role: ["Reader"] },
    { user: "tim@nudecision.com", role: ["Administrator"] },
    { user: "tim@mtjohnston.net", role: ["Reader"] },
  ];

  const userroles = RoleMapping.filter(function (rl) {
    return rl.user.toLowerCase() == name1.toLowerCase();
  });

  const responseMessage = userroles.length > 0 ? userroles[0].role : ["no roles assigned"];

  //const responseMessage="found1 "+name;
  context.res = {
    // status: 200, /* Defaults to 200 */
    // body: "found1 "+name1
    body: responseMessage,
  };
};
export default httpTrigger;
