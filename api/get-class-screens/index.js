const { TableClient } = require("@azure/data-tables");
const { BlobServiceClient, BlobSASPermissions, SASProtocol } = require("@azure/storage-blob");
const { getEmail, isTeacher } = require("../checkMember");
const { setJson, setErrorJson } = require("../contextHelper");


const chatStorageAccountConnectionString = process.env.chatStorageAccountConnectionString;

const blobServiceClient = BlobServiceClient.fromConnectionString(chatStorageAccountConnectionString);
const containerClient = blobServiceClient.getContainerClient("screen");
const classesTableClient = TableClient.fromConnectionString(chatStorageAccountConnectionString, "classes");


module.exports = async function (context, req) {
  const teacherEmail = getEmail(req);

  if (!await isTeacher(teacherEmail, context)) {
    setErrorJson(context, "Unauthorized", 401);
    return;
  }

  const classId = req.query.classId;

  const classIds = classId.split(",");
  const trimmedClassIds = classIds.map(id => id.trim());

  let entities = [];
  for (let i = 0; i < trimmedClassIds.length; i++) {
    const id = trimmedClassIds[i];
    let continuationToken = null;
    let pageEntities = undefined;

    do {
      const page = await classesTableClient.listEntities({
        queryOptions: {
          filter: `PartitionKey eq '${id}'`
        }
      }).byPage({ maxPageSize: 100, continuationToken: continuationToken }).next();
      pageEntities = page.value;
      continuationToken = pageEntities.continuationToken;
      entities = entities.concat(pageEntities);
    }
    while (continuationToken !== undefined);
  }

  let screens = await Promise.all(entities.map(async entity => {
    const studentEmail = entity.rowKey;
    context.log(studentEmail);
    const blobName = studentEmail.replace(/[^a-zA-Z0-9 ]/g, '_') + ".jpeg";
    const blobClient = containerClient.getBlobClient(blobName);
    const sasUrl = await blobClient.generateSasUrl({
      protocol: SASProtocol.Https,
      permissions: BlobSASPermissions.parse("r"),
      startsOn: new Date(),
      expiresOn: new Date(new Date().valueOf() + (1 * 60 * 1000))
    });
    context.log(sasUrl);
    return { email: studentEmail, sasUrl, name: entity.Name };
  }));

  screens = screens.sort((p1, p2) => p1.name.localeCompare(p2.name));
  setJson(context, screens);
}
