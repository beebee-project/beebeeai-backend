const User = require("../models/User");
const { deleteObject } = require("../utils/storage");

async function resetUserToFreeState(userId, subscriptionPatch = {}) {
  const user = await User.findById(userId).select(
    "plan usage uploadedFiles subscription",
  );
  if (!user) return { ok: false, reason: "USER_NOT_FOUND" };

  for (const f of user.uploadedFiles || []) {
    const name = f.localName || f.gcsName;
    if (!name) continue;

    try {
      await deleteObject(name);
    } catch (e) {
      console.warn(
        "[resetUserToFreeState] file delete failed:",
        name,
        e?.message,
      );
    }
  }

  user.plan = "FREE";
  user.uploadedFiles = [];
  user.usage = {
    formulaConversions: 0,
    fileUploads: 0,
    lastReset: new Date(),
  };

  user.subscription = {
    ...(user.subscription || {}),
    ...subscriptionPatch,
  };

  await user.save();

  return { ok: true };
}

module.exports = { resetUserToFreeState };
