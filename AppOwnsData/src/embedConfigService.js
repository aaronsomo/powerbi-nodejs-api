const auth = require(__dirname + "/authentication.js");
const config = require(__dirname + "/../config/config.json");
const utils = require(__dirname + "/utils.js");
const PowerBiReportDetails = require(__dirname +
  "/../models/embedReportConfig.js");
const EmbedConfig = require(__dirname + "/../models/embedConfig.js");
const fetch = require("node-fetch");

async function getEmbedInfo() {
  try {
    const embedParams = await getEmbedParamsForSingleReport(
      config.workspaceId,
      config.reportId
    );

    return {
      accessToken: embedParams.embedToken.token,
      embedUrl: embedParams.reportsDetail,
      expiry: embedParams.embedToken.expiration,
      status: 200,
    };
  } catch (err) {
    return {
      status: err.status,
      error: `Error while retrieving report embed details\r\n${
        err.statusText
      }\r\nRequestId: \n${err.headers.get("requestid")}`,
    };
  }
}

async function getEmbedParamsForSingleReport(
  workspaceId,
  reportId,
  additionalDatasetId
) {
  const reportInGroupApi = `https://api.powerbi.com/v1.0/myorg/groups/${workspaceId}/reports/${reportId}`;
  const headers = await getRequestHeader();

  const result = await fetch(reportInGroupApi, {
    method: "GET",
    headers: headers,
  });

  if (!result.ok) {
    throw result;
  }

  const resultJson = await result.json();

  const reportDetails = new PowerBiReportDetails(
    resultJson.id,
    resultJson.name,
    resultJson.embedUrl
  );
  const reportEmbedConfig = new EmbedConfig();

  reportEmbedConfig.reportsDetail = [reportDetails];

  let datasetIds = [resultJson.datasetId];

  if (additionalDatasetId) {
    datasetIds.push(additionalDatasetId);
  }

  reportEmbedConfig.embedToken =
    await getEmbedTokenForSingleReportSingleWorkspace(
      reportId,
      datasetIds,
      workspaceId
    );
  return reportEmbedConfig;
}

async function getEmbedParamsForMultipleReports(
  workspaceId,
  reportIds,
  additionalDatasetIds
) {
  const reportEmbedConfig = new EmbedConfig();

  reportEmbedConfig.reportsDetail = [];

  let datasetIds = [];

  for (const reportId of reportIds) {
    const reportInGroupApi = `https://api.powerbi.com/v1.0/myorg/groups/${workspaceId}/reports/${reportId}`;
    const headers = await getRequestHeader();

    const result = await fetch(reportInGroupApi, {
      method: "GET",
      headers: headers,
    });

    if (!result.ok) {
      throw result;
    }

    const resultJson = await result.json();

    const reportDetails = new PowerBiReportDetails(
      resultJson.id,
      resultJson.name,
      resultJson.embedUrl
    );

    reportEmbedConfig.reportsDetail.push(reportDetails);

    datasetIds.push(resultJson.datasetId);
  }

  if (additionalDatasetIds) {
    datasetIds.push(...additionalDatasetIds);
  }

  reportEmbedConfig.embedToken =
    await getEmbedTokenForMultipleReportsSingleWorkspace(
      reportIds,
      datasetIds,
      workspaceId
    );
  return reportEmbedConfig;
}

async function getEmbedTokenForSingleReportSingleWorkspace(
  reportId,
  datasetIds,
  targetWorkspaceId
) {
  let formData = {
    reports: [
      {
        id: reportId,
      },
    ],
  };

  formData["datasets"] = [];
  for (const datasetId of datasetIds) {
    formData["datasets"].push({
      id: datasetId,
    });
  }

  if (targetWorkspaceId) {
    formData["targetWorkspaces"] = [];
    formData["targetWorkspaces"].push({
      id: targetWorkspaceId,
    });
  }

  const embedTokenApi = "https://api.powerbi.com/v1.0/myorg/GenerateToken";
  const headers = await getRequestHeader();

  const result = await fetch(embedTokenApi, {
    method: "POST",
    headers: headers,
    body: JSON.stringify(formData),
  });

  if (!result.ok) throw result;
  return result.json();
}

async function getEmbedTokenForMultipleReportsSingleWorkspace(
  reportIds,
  datasetIds,
  targetWorkspaceId
) {
  let formData = { datasets: [] };
  for (const datasetId of datasetIds) {
    formData["datasets"].push({
      id: datasetId,
    });
  }

  formData["reports"] = [];
  for (const reportId of reportIds) {
    formData["reports"].push({
      id: reportId,
    });
  }

  if (targetWorkspaceId) {
    formData["targetWorkspaces"] = [];
    formData["targetWorkspaces"].push({
      id: targetWorkspaceId,
    });
  }

  const embedTokenApi = "https://api.powerbi.com/v1.0/myorg/GenerateToken";
  const headers = await getRequestHeader();

  const result = await fetch(embedTokenApi, {
    method: "POST",
    headers: headers,
    body: JSON.stringify(formData),
  });

  if (!result.ok) throw result;
  return result.json();
}

async function getEmbedTokenForMultipleReportsMultipleWorkspaces(
  reportIds,
  datasetIds,
  targetWorkspaceIds
) {
  let formData = { datasets: [] };
  for (const datasetId of datasetIds) {
    formData["datasets"].push({
      id: datasetId,
    });
  }

  formData["reports"] = [];
  for (const reportId of reportIds) {
    formData["reports"].push({
      id: reportId,
    });
  }

  if (targetWorkspaceIds) {
    formData["targetWorkspaces"] = [];
    for (const targetWorkspaceId of targetWorkspaceIds) {
      formData["targetWorkspaces"].push({
        id: targetWorkspaceId,
      });
    }
  }

  const embedTokenApi = "https://api.powerbi.com/v1.0/myorg/GenerateToken";
  const headers = await getRequestHeader();

  const result = await fetch(embedTokenApi, {
    method: "POST",
    headers: headers,
    body: JSON.stringify(formData),
  });

  if (!result.ok) throw result;
  return result.json();
}

async function getRequestHeader() {
  let tokenResponse;

  let errorResponse;

  try {
    tokenResponse = await auth.getAccessToken();
  } catch (err) {
    if (
      err.hasOwnProperty("error_description") &&
      err.hasOwnProperty("error")
    ) {
      errorResponse = err.error_description;
    } else {
      errorResponse = err.toString();
    }
    return {
      status: 401,
      error: errorResponse,
    };
  }

  const token = tokenResponse.accessToken;
  return {
    "Content-Type": "application/json",
    Authorization: utils.getAuthHeader(token),
  };
}

module.exports = {
  getEmbedInfo: getEmbedInfo,
};
