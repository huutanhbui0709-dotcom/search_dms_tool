const https = require("https");
const crypto = require("crypto");

const R2_ACCOUNT_ID        = process.env.R2_ACCOUNT_ID        || "";
const R2_ACCESS_KEY_ID     = process.env.R2_ACCESS_KEY_ID     || "";
const R2_SECRET_ACCESS_KEY = process.env.R2_SECRET_ACCESS_KEY || "";
const R2_BUCKET_NAME       = process.env.R2_BUCKET_NAME       || "";
const R2_OBJECT_KEY        = "task_pending_data.json";

function hmacSha256(key, msg, enc) {
  return crypto.createHmac("sha256", key).update(msg, "utf8").digest(enc);
}
function sha256hex(msg) {
  return crypto.createHash("sha256").update(msg, "utf8").digest("hex");
}

function buildR2Request(method, body) {
  body = body || "";
  const host      = R2_ACCOUNT_ID + ".r2.cloudflarestorage.com";
  const uriPath   = "/" + R2_BUCKET_NAME + "/" + R2_OBJECT_KEY;
  const now       = new Date();
  const dateStamp = now.toISOString().slice(0, 10).replace(/-/g, "");
  const amzDate   = now.toISOString().replace(/[:-]/g, "").replace(/\.\d+/, "");
  const region    = "auto";
  const service   = "s3";

  const payloadHash  = sha256hex(body);
  const contentType  = method === "PUT" ? "application/json" : "";

  let canonicalHeaders = "host:" + host + "\nx-amz-content-sha256:" + payloadHash + "\nx-amz-date:" + amzDate + "\n";
  let signedHeaders    = "host;x-amz-content-sha256;x-amz-date";
  if (contentType) {
    canonicalHeaders = "content-type:" + contentType + "\n" + canonicalHeaders;
    signedHeaders    = "content-type;" + signedHeaders;
  }

  const canonicalRequest = [method, uriPath, "", canonicalHeaders, signedHeaders, payloadHash].join("\n");
  const credentialScope  = dateStamp + "/" + region + "/" + service + "/aws4_request";
  const stringToSign     = ["AWS4-HMAC-SHA256", amzDate, credentialScope, sha256hex(canonicalRequest)].join("\n");

  const kDate    = hmacSha256("AWS4" + R2_SECRET_ACCESS_KEY, dateStamp);
  const kRegion  = hmacSha256(kDate, region);
  const kService = hmacSha256(kRegion, service);
  const kSigning = hmacSha256(kService, "aws4_request");
  const signature = hmacSha256(kSigning, stringToSign, "hex");

  const auth = "AWS4-HMAC-SHA256 Credential=" + R2_ACCESS_KEY_ID + "/" + credentialScope + ",SignedHeaders=" + signedHeaders + ",Signature=" + signature;

  const headers = {
    "Authorization":        auth,
    "x-amz-date":           amzDate,
    "x-amz-content-sha256": payloadHash
  };
  if (contentType) headers["Content-Type"] = contentType;
  return { host: host, path: uriPath, headers: headers };
}

function r2Fetch(method, body) {
  body = body || "";
  return new Promise(function(resolve, reject) {
    const info = buildR2Request(method, body);
    const options = { hostname: info.host, path: info.path, method: method, headers: info.headers };
    const req = https.request(options, function(res) {
      let data = "";
      res.on("data", function(c) { data += c; });
      res.on("end", function() { resolve({ statusCode: res.statusCode, body: data }); });
    });
    req.on("error", reject);
    if (body) req.write(body);
    req.end();
  });
}

const CORS = {
  "Access-Control-Allow-Origin":  "*",
  "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type"
};

// Đọc dữ liệu hiện tại từ R2, fallback về object trống nếu chưa có hoặc lỗi
async function getTasksData() {
  const getR = await r2Fetch("GET", "");
  if (getR.statusCode === 200) {
    try {
      const data = JSON.parse(getR.body);
      return {
        tasks: data.tasks || [],
        members: data.members || []
      };
    } catch (e) {
      return { tasks: [], members: [] };
    }
  }
  return { tasks: [], members: [] };
}

// Lưu dữ liệu vào R2
async function saveTasksData(data) {
  const putBody = JSON.stringify(data);
  const putR = await r2Fetch("PUT", putBody);
  if (putR.statusCode >= 200 && putR.statusCode < 300) {
    return { status: "success", data };
  } else {
    throw new Error("R2 PUT failed HTTP " + putR.statusCode + ": " + putR.body);
  }
}

module.exports = async function handler(req, res) {
  // CORS preflight
  Object.entries(CORS).forEach(([k, v]) => res.setHeader(k, v));
  if (req.method === "OPTIONS") return res.status(204).end();

  try {
    if (req.method === "GET") {
      const urlParts = require("url").parse(req.url, true);
      const action = urlParts.query.action;

      const data = await getTasksData();

      if (action === "read_tasks") {
        return res.status(200).json({
          status: "success",
          data: data.tasks,
          tasks: data.tasks
        });
      } else if (action === "read_members") {
        return res.status(200).json({
          status: "success",
          data: data.members
        });
      } else {
        // Trả về toàn bộ dữ liệu mặc định
        return res.status(200).json({
          status: "success",
          tasks: data.tasks,
          members: data.members
        });
      }
    }

    if (req.method === "POST") {
      const payload = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
      const { action } = payload;
      
      const currentData = await getTasksData();

      if (action === "save_task") {
        const { id, deadline, issue, note, assignees, status, priority } = payload;
        const taskStatus = (String(status || "").trim().toLowerCase() === "done") ? "Done" : "Pending";
        const taskPriority = ["P1","P2","P3"].includes(priority) ? priority : "P3";
        
        if (id) {
          // Edit task
          const taskIndex = currentData.tasks.findIndex(t => Number(t.id) === Number(id));
          if (taskIndex !== -1) {
            currentData.tasks[taskIndex] = {
              id: Number(id),
              deadline: deadline || "",
              issue: issue || "",
              note: note || "",
              assignees: Array.isArray(assignees) ? assignees : [],
              status: taskStatus,
              priority: taskPriority
            };
          } else {
            return res.status(404).json({ status: "error", message: "Không tìm thấy task cần chỉnh sửa" });
          }
        } else {
          // Add task new
          const newTask = {
            id: Date.now(), // ID dựa trên timestamp
            deadline: deadline || "",
            issue: issue || "",
            note: note || "",
            assignees: Array.isArray(assignees) ? assignees : [],
            status: taskStatus,
            priority: taskPriority
          };
          currentData.tasks.push(newTask);
        }

        await saveTasksData(currentData);
        return res.status(200).json({ status: "success" });
      }

      if (action === "delete_task") {
        const { id } = payload;
        currentData.tasks = currentData.tasks.filter(t => Number(t.id) !== Number(id));
        await saveTasksData(currentData);
        return res.status(200).json({ status: "success" });
      }

      if (action === "add_member") {
        const name = (payload.name || "").trim();
        if (!name) return res.status(400).json({ status: "error", message: "Tên không được để trống" });
        
        const exists = currentData.members.some(m => m.name.toLowerCase() === name.toLowerCase());
        if (exists) return res.status(400).json({ status: "error", message: "Thành viên đã tồn tại" });
        
        currentData.members.push({ name });
        await saveTasksData(currentData);
        return res.status(200).json({ status: "success" });
      }

      if (action === "delete_member") {
        const name = (payload.name || "").trim();
        if (!name) return res.status(400).json({ status: "error", message: "Tên không được để trống" });
        
        const initialLen = currentData.members.length;
        currentData.members = currentData.members.filter(m => m.name.toLowerCase() !== name.toLowerCase());
        
        if (currentData.members.length === initialLen) {
          return res.status(404).json({ status: "error", message: "Không tìm thấy thành viên: " + name });
        }
        
        await saveTasksData(currentData);
        return res.status(200).json({ status: "success" });
      }

      return res.status(400).json({ status: "error", message: "Action POST không được hỗ trợ" });
    }

    return res.status(405).end("Method Not Allowed");
  } catch (err) {
    return res.status(500).json({ status: "error", message: err.message, stack: err.stack });
  }
};
