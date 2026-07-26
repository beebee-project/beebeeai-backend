const jwt = require("jsonwebtoken");
const User = require("../models/User");

const AUTH_ERROR_CODES = Object.freeze({
  TOKEN_MISSING: "TOKEN_MISSING",
  TOKEN_FORMAT_INVALID: "TOKEN_FORMAT_INVALID",
  TOKEN_EXPIRED: "TOKEN_EXPIRED",
  TOKEN_INVALID: "TOKEN_INVALID",
  TOKEN_NOT_ACTIVE: "TOKEN_NOT_ACTIVE",
  USER_NOT_FOUND: "AUTH_USER_NOT_FOUND",
  CONFIG_ERROR: "AUTH_CONFIG_ERROR",
});

function authorizationHeader(req = {}) {
  return String(req.headers?.authorization || "").trim();
}

function extractBearerToken(req = {}) {
  const header = authorizationHeader(req);
  const match = header.match(/^Bearer\s+([^\s]+)$/i);
  return match ? match[1] : "";
}

function setNoStoreHeaders(res) {
  if (typeof res.set === "function") {
    res.set("Cache-Control", "no-store");
    res.set("Pragma", "no-cache");
  } else if (typeof res.setHeader === "function") {
    res.setHeader("Cache-Control", "no-store");
    res.setHeader("Pragma", "no-cache");
  }
}

function sendAuthError(
  res,
  {
    status = 401,
    code,
    message,
    expiredAt = null,
    authenticateError = "invalid_token",
    authenticateDescription = "authentication failed",
  } = {},
) {
  setNoStoreHeaders(res);

  const challenge = [
    "Bearer",
    `error=\"${authenticateError}\"`,
    `error_description=\"${authenticateDescription}\"`,
  ].join(", ");

  if (typeof res.set === "function") {
    res.set("WWW-Authenticate", challenge);
  } else if (typeof res.setHeader === "function") {
    res.setHeader("WWW-Authenticate", challenge);
  }

  const body = {
    ok: false,
    code,
    message,
    reauthRequired: status === 401,
  };

  if (expiredAt) {
    body.expiredAt = new Date(expiredAt).toISOString();
  }

  return res.status(status).json(body);
}

function localDevUser() {
  return {
    id: "000000000000000000000001",
    email: "dev@local.test",
    plan: "PRO",
    usage: {
      templateGenerations: 0,
      fileUploads: 0,
    },
    uploadedFiles: [],
  };
}

const protect = async (req, res, next) => {
  if (process.env.LOCAL_DEV === "1" && process.env.DEV_BYPASS_AUTH === "1") {
    req.user = localDevUser();
    return next();
  }

  const header = authorizationHeader(req);
  if (!header) {
    return sendAuthError(res, {
      code: AUTH_ERROR_CODES.TOKEN_MISSING,
      message: "인증 실패: 토큰이 없습니다.",
      authenticateError: "invalid_request",
      authenticateDescription: "bearer token is missing",
    });
  }

  const token = extractBearerToken(req);
  if (!token) {
    return sendAuthError(res, {
      code: AUTH_ERROR_CODES.TOKEN_FORMAT_INVALID,
      message: "인증 실패: Authorization 헤더 형식이 올바르지 않습니다.",
      authenticateError: "invalid_request",
      authenticateDescription: "authorization header must use Bearer token",
    });
  }

  const secret = String(process.env.JWT_SECRET || "").trim();
  if (!secret) {
    console.error("[auth] JWT_SECRET is not configured.");
    return sendAuthError(res, {
      status: 500,
      code: AUTH_ERROR_CODES.CONFIG_ERROR,
      message: "서버 인증 설정 오류가 발생했습니다.",
      authenticateError: "server_error",
      authenticateDescription: "jwt secret is not configured",
    });
  }

  let decoded;
  try {
    decoded = jwt.verify(token, secret);
  } catch (error) {
    if (error?.name === "TokenExpiredError") {
      return sendAuthError(res, {
        code: AUTH_ERROR_CODES.TOKEN_EXPIRED,
        message: "로그인 세션이 만료되었습니다. 다시 로그인해주세요.",
        expiredAt: error.expiredAt,
        authenticateDescription: "access token expired",
      });
    }

    if (error?.name === "NotBeforeError") {
      return sendAuthError(res, {
        code: AUTH_ERROR_CODES.TOKEN_NOT_ACTIVE,
        message: "아직 사용할 수 없는 인증 토큰입니다.",
        authenticateDescription: "access token is not active",
      });
    }

    if (error?.name !== "JsonWebTokenError") {
      console.error(
        "[auth] unexpected token verification error:",
        error?.message || error,
      );
    }

    return sendAuthError(res, {
      code: AUTH_ERROR_CODES.TOKEN_INVALID,
      message: "인증 실패: 유효하지 않은 토큰입니다.",
      authenticateDescription: "access token is invalid",
    });
  }

  if (!decoded?.id) {
    return sendAuthError(res, {
      code: AUTH_ERROR_CODES.TOKEN_INVALID,
      message: "인증 실패: 사용자 식별자가 없는 토큰입니다.",
      authenticateDescription: "access token has no user id",
    });
  }

  try {
    req.user = await User.findById(decoded.id).select("-password");
  } catch (error) {
    return next(error);
  }

  if (!req.user) {
    return sendAuthError(res, {
      code: AUTH_ERROR_CODES.USER_NOT_FOUND,
      message: "사용자를 찾을 수 없습니다.",
      authenticateDescription: "authenticated user no longer exists",
    });
  }

  req.auth = {
    tokenPayload: decoded,
  };

  return next();
};

module.exports = {
  AUTH_ERROR_CODES,
  extractBearerToken,
  protect,
};
