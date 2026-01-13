const mongoose = require("mongoose");
const User = require("./User");

async function main() {
  const uri = process.env.MONGO_URI;
  if (!uri) throw new Error("MONGO_URI is required");

  await mongoose.connect(uri);

  const filter = {
    $and: [
      {
        $or: [{ plan: "PRO" }, { "subscription.status": "ACTIVE" }],
      },
      {
        $or: [
          { "subscription.billingKey": { $exists: false } },
          { "subscription.billingKey": null },
          { "subscription.billingKey": "" },
        ],
      },
      {
        $or: [
          { "subscription.customerKey": { $exists: false } },
          { "subscription.customerKey": null },
          { "subscription.customerKey": "" },
        ],
      },
    ],
  };

  const update = {
    $set: {
      plan: "FREE",
      "subscription.status": "INACTIVE",
      "subscription.nextChargeAt": null,
      "subscription.lastChargeKey": null,
      "subscription.chargeLockKey": null,
    },
  };

  const r = await User.updateMany(filter, update);
  console.log("updated:", r.modifiedCount);

  await mongoose.disconnect();
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
