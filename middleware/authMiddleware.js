const jwt = require("jsonwebtoken");
const User = require("../models/User");

const protect = async (req, res, next) => {
  let token;

  if (
    req.headers.authorization &&
    req.headers.authorization.startsWith("Bearer")
  ) {
    try {
      token = req.headers.authorization.split(" ")[1];

      const decoded = jwt.verify(token, process.env.JWT_SECRET);

      req.user = await User.findById(decoded.id).select("-password");

      if (!req.user) {
        return res.status(401).json({ message: "사용자를 찾을 수 없습니다." });
      }

      next();
    } catch (error) {
      console.error(error);
      return res
        .status(401)
        .json({ message: "인증 실패: 유효하지 않은 토큰입니다." });
    }
  }

  if (!token) {
    return res.status(401).json({ message: "인증 실패: 토큰이 없습니다." });
  }
};

module.exports = { protect };
