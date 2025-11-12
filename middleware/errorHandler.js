const errorHandler = (err, req, res, next) => {
  console.error(err.stack);

  const statusCode = err.statusCode || 500;
  const message = err.message || "서버 내부 오류가 발생했습니다.";

  res.status(statusCode).json({
    status: "error",
    message: message,
  });
};

module.exports = errorHandler;
