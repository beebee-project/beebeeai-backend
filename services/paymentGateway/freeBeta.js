exports.createCheckoutSession = async ({ userId, plan }) => {
  return {
    orderId: `free_${userId}_${Date.now()}`,
    orderName: plan === "FREE" ? "무료 플랜" : "테스트 플랜",
    amount: 0,
    provider: "freeBeta",
  };
};

exports.confirmPayment = async () => {
  return { status: "FREE", message: "결제 불필요 - 무료 플랜" };
};

exports.cancelPayment = async () => {
  return { status: "FREE", message: "무료 플랜은 취소 불필요" };
};

exports.parseAndVerifyWebhook = async () => {
  return { eventType: "freeBeta.noop", data: {} };
};
