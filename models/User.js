const mongoose = require("mongoose");
const bcrypt = require("bcryptjs");
const crypto = require("crypto");

const fileSchema = new mongoose.Schema({
  originalName: { type: String, required: true },
  gcsName: { type: String, required: true },
  size: { type: Number, required: true },
  uploadDate: { type: Date, default: Date.now },
});

const userSchema = new mongoose.Schema({
  email: {
    type: String,
    required: [true, "이메일을 입력해주세요."],
    unique: true,
    lowercase: true,
  },
  password: {
    type: String,

    minlength: 6,
    select: false,
  },

  name: {
    type: String,
    default: "사용자",
  },
  googleId: {
    type: String,
  },
  isVerified: {
    type: Boolean,
    default: false,
  },
  emailVerificationToken: String,
  emailVerificationExpires: Date,
  passwordResetToken: String,
  passwordResetExpires: Date,
  uploadedFiles: [fileSchema],
  usage: {
    formulaConversions: { type: Number, default: 0 },
    fileUploads: { type: Number, default: 0 },
    lastReset: { type: Date, default: Date.now },
  },
  plan: { type: String, enum: ["FREE", "PRO"], default: "FREE" },
  subscription: {
    customerKey: { type: String },
    billingKey: { type: String },
    status: {
      type: String,
      enum: [
        "NONE",
        "ACTIVE",
        "PAST_DUE",
        "CANCELED_PENDING",
        "CANCELED",
        "INACTIVE",
      ],
      default: "INACTIVE",
    },
    // ✅ 중복 결제 방지/운영 안정성용 필드들
    lastChargeKey: { type: String },
    lastOrderId: { type: String },
    lastChargeAttemptAt: { type: Date },
    lastChargeError: { type: String },

    // ✅ 동시 실행(중복 호출) 방지용 락
    chargeLockKey: { type: String },

    startedAt: { type: Date },
    nextChargeAt: { type: Date },
    lastChargedAt: { type: Date },
    cancelAtPeriodEnd: { type: Boolean },
    canceledAt: { type: Date },
    endedAt: { type: Date },
    lastPaymentKey: { type: String },

    isDeleted: { type: Boolean, default: false, index: true },
    deletedAt: { type: Date, default: null },
    purgeAt: { type: Date, default: null, index: true },

    authIdentity: {
      emailHash: { type: String, default: null, index: true },
      googleId: { type: String, default: null, index: true },
    },
  },
});

// 비밀번호가 존재하고 변경되었을 때만 암호화
userSchema.pre("save", async function (next) {
  if (!this.isModified("password") || !this.password) return next();
  const salt = await bcrypt.genSalt(10);
  this.password = await bcrypt.hash(this.password, salt);
  next();
});

userSchema.methods.comparePassword = async function (candidatePassword) {
  return await bcrypt.compare(candidatePassword, this.password);
};

userSchema.methods.createEmailVerificationToken = function () {
  const verificationToken = crypto.randomBytes(32).toString("hex");
  this.emailVerificationToken = crypto
    .createHash("sha256")
    .update(verificationToken)
    .digest("hex");
  this.emailVerificationExpires = Date.now() + 10 * 60 * 1000;
  return verificationToken;
};

userSchema.methods.createPasswordResetToken = function () {
  const resetToken = crypto.randomBytes(32).toString("hex");
  this.passwordResetToken = crypto
    .createHash("sha256")
    .update(resetToken)
    .digest("hex");
  this.passwordResetExpires = Date.now() + 10 * 60 * 1000;
  return resetToken;
};

const User = mongoose.model("User", userSchema);
module.exports = User;
