import { PublicClientApplication, Configuration } from "@azure/msal-browser";
import { RowData, validateRow, getHoleSizeFromSection } from "../types";

const msalConfig: Configuration = {
  auth: {
    clientId: "557f78ec-a103-4736-b25b-c5672e3cd2fc",
    authority: "https://login.microsoftonline.com/b7be7686-6f97-4db7-9081-a23cf09a96b5",
    redirectUri: typeof window !== "undefined" ? window.location.origin : undefined,
  },
  cache: {
    cacheLocation: "sessionStorage",
  },
};

export const msalInstance = new PublicClientApplication(msalConfig);

let isMsalInitialized = false;

async function ensureMsalInit() {
  if (!isMsalInitialized) {
    if (typeof msalInstance.initialize === "function") {
      await msalInstance.initialize();
    }
    isMsalInitialized = true;
  }
}

export const loginRequest = {
  scopes: ["Files.Read"],
};

/**
 * Initiates the sign-in popup flow via MSAL.
 * Sets the active account and returns account + access token.
 */
export const msSignIn = async (): Promise<{ account: any; accessToken: string }> => {
  await ensureMsalInit();
  const response = await msalInstance.loginPopup(loginRequest);
  if (response && response.account) {
    msalInstance.setActiveAccount(response.account);
    console.log("Signed in as:", response.account.username);
    try {
      await fetchExcelData();
    } catch (e) {
      console.warn("Auto fetchExcelData on sign-in warning:", e);
    }
  }
  return {
    account: response.account,
    accessToken: response.accessToken,
  };
};

export const signIn = msSignIn;

/**
 * Retrieves a valid access token for the active account via acquireTokenSilent or acquireTokenPopup fallback.
 */
export const getAccessToken = async (): Promise<string> => {
  await ensureMsalInit();
  let account = msalInstance.getActiveAccount();
  if (!account) {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      account = accounts[0];
      msalInstance.setActiveAccount(account);
    }
  }
  if (!account) throw new Error("No signed-in Microsoft account");

  try {
    const result = await msalInstance.acquireTokenSilent({
      ...loginRequest,
      account,
    });
    return result.accessToken;
  } catch (err) {
    const result = await msalInstance.acquireTokenPopup(loginRequest);
    return result.accessToken;
  }
};

/**
 * Converts Excel date serial number to JavaScript Date object.
 */
export function excelSerialToDate(serial: any): Date | null {
  if (!serial || typeof serial !== "number") return null;
  const excelEpoch = new Date(Date.UTC(1899, 11, 30));
  const msPerDay = 24 * 60 * 60 * 1000;
  return new Date(excelEpoch.getTime() + serial * msPerDay);
}

/**
 * Fetches Excel table rows from Microsoft Graph API workbook tables.
 */
export async function fetchExcelData(): Promise<Record<string, any>[] | null> {
  const token = await getAccessToken();

  const driveId = "b!ZquY72zAo0SLlAmzWa9v9uzZ5eLahXtKrzaCWCvisj8EZeSVnOaIQ7ieKyHEWfBt";
  const fileId = "01AT5YVTHU3SJ2NAO2XNB3GAH4N7PFTSQN";
  const tableName = "Table101213";

  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}/workbook/tables('${tableName}')/rows`;

  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!response.ok) {
    const errText = await response.text();
    console.error("Graph fetch failed:", errText);
    return null;
  }

  const data = await response.json();

  // Column order matches the actual sheet layout, left to right
  const columns = [
    "Client", "Rig", "Well", "Section", "Qty", "Configuration",
    "Footage", "Date", "Month", "ShopDate", "Comments",
    "Column1", "Column2", "Column3"
  ];

  const rows = (data.value || [])
    .map((r: any) => r.values[0])
    .filter((row: any[]) => row && row.some((cell: any) => cell !== "" && cell !== null && cell !== undefined))
    .map((rowArray: any[]) => {
      const obj: Record<string, any> = {};
      columns.forEach((colName, i) => {
        let value = rowArray[i];
        if ((colName === "Date" || colName === "Month") && typeof value === "number") {
          value = excelSerialToDate(value);
        }
        obj[colName] = value;
      });
      return obj;
    });

  console.log(`Fetched ${rows.length} rows`);
  return rows;
}

/**
  * Converts rows returned by fetchExcelData into formatted dashboard RowData array.
  */
export function convertDispatchRowsToRowData(dispatchRows: any[]): RowData[] {
  if (!Array.isArray(dispatchRows)) return [];

  return dispatchRows.map((r, index) => {
    // Format Date
    let dateStr = "";
    if (r.Date) {
      if (r.Date instanceof Date) {
        const d = r.Date;
        const day = String(d.getUTCDate()).padStart(2, "0");
        const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        const month = months[d.getUTCMonth()] || "";
        const year = String(d.getUTCFullYear()).slice(-2);
        dateStr = `${day}-${month}-${year}`;
      } else {
        dateStr = String(r.Date).trim();
      }
    }

    // Format Month
    let monthStr = "";
    if (r.Month) {
      if (r.Month instanceof Date) {
        const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        monthStr = months[r.Month.getUTCMonth()] || "";
      } else {
        monthStr = String(r.Month).trim();
      }
    }

    const qtyNum = Number(r.Qty);
    const qty = isNaN(qtyNum) ? 1 : qtyNum;

    const footageNum = Number(r.Footage);
    const footage = (r.Footage !== null && r.Footage !== undefined && r.Footage !== "" && !isNaN(footageNum))
      ? footageNum
      : null;

    const sectionStr = String(r.Section || "").trim();
    const holeSizeStr = String(r.HoleSize || r.holeSize || getHoleSizeFromSection(sectionStr)).trim();

    const rowObj: RowData = {
      id: `ms-row-${index}-${Date.now()}`,
      client: String(r.Client || "").trim(),
      rig: String(r.Rig || "").trim(),
      well: String(r.Well || "").trim(),
      section: sectionStr,
      holeSize: holeSizeStr,
      qty,
      configuration: String(r.Configuration || "").trim(),
      footage,
      date: dateStr,
      month: monthStr,
      shopDate: String(r.ShopDate || "").trim(),
      comments: String(r.Comments || "").trim()
    };

    rowObj.validationErrors = validateRow(rowObj);
    return rowObj;
  });
}

/**
 * Signs out the current active MSAL account.
 */
export const msSignOut = async (): Promise<void> => {
  try {
    await ensureMsalInit();
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      const logoutRequest = {
        account: accounts[0],
        postLogoutRedirectUri: window.location.origin,
      };
      if (typeof msalInstance.logoutPopup === "function") {
        await msalInstance.logoutPopup(logoutRequest);
      } else {
        await msalInstance.logoutRedirect(logoutRequest);
      }
    }
  } catch (error) {
    console.error("Error signing out via MSAL:", error);
  }
};

/**
 * Searches OneDrive for .xlsx files.
 */
export const searchOneDriveExcelFiles = async (accessToken: string): Promise<any[]> => {
  try {
    const url = "https://graph.microsoft.com/v1.0/me/drive/root/search(q='.xlsx')";
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`OneDrive API error: ${response.status} - ${errText}`);
    }

    const data = await response.json();
    const files = (data.value || []).map((item: any) => ({
      id: item.id,
      name: item.name,
      lastModifiedDateTime: item.lastModifiedDateTime || item.fileSystemInfo?.lastModifiedDateTime,
    }));
    return files;
  } catch (error) {
    console.error("Error searching OneDrive Excel files:", error);
    throw error;
  }
};

/**
 * Downloads the content of a OneDrive file as an ArrayBuffer.
 */
export const fetchOneDriveFileContent = async (
  accessToken: string,
  fileId: string
): Promise<ArrayBuffer> => {
  try {
    const url = `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/content`;
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    });

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`OneDrive Download error: ${response.status} - ${errText}`);
    }

    return await response.arrayBuffer();
  } catch (error) {
    console.error("Error fetching OneDrive file content:", error);
    throw error;
  }
};
