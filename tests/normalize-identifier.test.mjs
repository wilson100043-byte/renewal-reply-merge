import assert from "node:assert/strict";
import test from "node:test";

import { normalizeIdentifier, resolveRenewalStatus } from "../xlsx-engine.js";

test("normalizes Excel scientific notation without changing text identifiers", () => {
  assert.equal(normalizeIdentifier("1.140904601E9"), "1140904601");
  assert.equal(normalizeIdentifier("1.140604502e+9"), "1140604502");
  assert.equal(normalizeIdentifier("1140904601.0"), "1140904601");
  assert.equal(normalizeIdentifier("001140904601"), "001140904601");
  assert.equal(normalizeIdentifier("customer-01"), "customer-01");
  assert.equal(normalizeIdentifier("1e21"), "1e21");
});

test("keeps a local OK renewal status while other statuses follow the source", () => {
  assert.equal(resolveRenewalStatus("續訂", "OK"), "OK");
  assert.equal(resolveRenewalStatus("停訂", " ok "), "ok");
  assert.equal(resolveRenewalStatus("續訂", "停訂"), "續訂");
  assert.equal(resolveRenewalStatus("", "停訂"), "");
});
