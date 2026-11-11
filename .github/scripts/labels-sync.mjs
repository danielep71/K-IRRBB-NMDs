#!/usr/bin/env node

import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { appendFile, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { createServer } from "node:http";
import { tmpdir } from "node:os";
import { join } from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import { inlineCode, markdownEscape, selfTestMarkdown } from "./markdown.mjs";

const EXPECTED_SCHEMA_VERSION = 1;
const DEFAULT_MANIFEST = ".github/labels.json";
const DEFAULT_POLICY = ".github/label-policy.json";
const API_VERSION = "2022-11-28";
const DEFAULT_API_URL = "https://api.github.com";
const PROFILE_NAMES = ["application", "library", "ui-component"];
const LABELS_PER_PAGE = 100;
const MAX_LABEL_PAGES = 100;
const ISSUES_PER_PAGE = 100;
const MAX_ISSUE_PAGES = 100;
const USAGE_SAMPLE = 10;
const MAX_RATE_LIMIT_RETRIES = 2;
const MAX_RETRY_DELAY_MS = 5000;

function compareNames(left, right) {
  const leftKey = left.toLowerCase();
  const rightKey = right.toLowerCase();
  if (leftKey < rightKey) return -1;
  if (leftKey > rightKey) return 1;
  return left < right ? -1 : left > right ? 1 : 0;
}

function sameKeys(value, expected) {
  const actual = Object.keys(value).sort();
  return JSON.stringify(actual) === JSON.stringify([...expected].sort());
}

function normalizedLiveLabel(label) {
  return {
    name: String(label.name ?? ""),
    color: String(label.color ?? "").toUpperCase(),
    description: label.description == null ? "" : String(label.description)
  };
}

function validateLabelArray(labels, location, seen) {
  const errors = [];
  if (!Array.isArray(labels)) return [`${location} must be an array`];

  labels.forEach((label, index) => {
    const item = `${location}[${index}]`;
    if (!label || typeof label !== "object" || Array.isArray(label)) {
      errors.push(`${item} must be an object`);
      return;
    }
    if (!sameKeys(label, ["name", "color", "description"])) {
      errors.push(`${item} must contain exactly name, color, and description`);
    }

    if (typeof label.name !== "string") {
      errors.push(`${item}.name must be a string`);
    } else {
      if (label.name.length === 0 || label.name !== label.name.trim()) {
        errors.push(`${item}.name must be non-empty with no surrounding whitespace`);
      }
      if (label.name.length > 50) errors.push(`${item}.name exceeds GitHub's 50-character limit`);
      if (/\r|\n/.test(label.name)) errors.push(`${item}.name must be single-line`);
      const key = label.name.toLowerCase();
      if (seen.has(key)) {
        errors.push(`${item}.name duplicates ${seen.get(key)} case-insensitively`);
      } else {
        seen.set(key, `${item}.name`);
      }
    }

    if (typeof label.color !== "string" || !/^[0-9A-F]{6}$/.test(label.color)) {
      errors.push(`${item}.color must be six uppercase hexadecimal characters without #`);
    }
    if (typeof label.description !== "string") {
      errors.push(`${item}.description must be a string`);
    } else {
      if (label.description.length === 0) errors.push(`${item}.description must not be empty`);
      if (label.description.length > 100) errors.push(`${item}.description exceeds GitHub's 100-character limit`);
      if (/\r|\n/.test(label.description)) errors.push(`${item}.description must be single-line`);
    }
  });

  for (let index = 1; index < labels.length; index += 1) {
    const previous = labels[index - 1]?.name;
    const current = labels[index]?.name;
    if (typeof previous === "string" && typeof current === "string" && compareNames(previous, current) >= 0) {
      errors.push(`${location} must be sorted case-insensitively by name: ${previous} precedes ${current}`);
      break;
    }
  }
  return errors;
}

export function validateManifest(document) {
  const errors = [];
  if (!document || typeof document !== "object" || Array.isArray(document)) {
    throw new Error("Manifest must be a JSON object.");
  }
  if (!sameKeys(document, ["schema_version", "prune", "core", "overlays"])) {
    errors.push("root must contain exactly schema_version, prune, core, and overlays");
  }
  if (document.schema_version !== EXPECTED_SCHEMA_VERSION) {
    errors.push(`schema_version must be ${EXPECTED_SCHEMA_VERSION}`);
  }
  if (typeof document.prune !== "boolean") errors.push("prune must be a boolean");

  const seen = new Map();
  errors.push(...validateLabelArray(document.core, "core", seen));
  if (Array.isArray(document.core) && document.core.length === 0) errors.push("core must not be empty");

  const overlays = document.overlays;
  if (!overlays || typeof overlays !== "object" || Array.isArray(overlays)) {
    errors.push("overlays must be an object");
  } else if (!sameKeys(overlays, ["profile", "domain"])) {
    errors.push("overlays must contain exactly profile and domain");
  }

  const profiles = overlays?.profile;
  if (!profiles || typeof profiles !== "object" || Array.isArray(profiles)) {
    errors.push("overlays.profile must be an object");
  } else {
    if (!sameKeys(profiles, PROFILE_NAMES)) {
      errors.push(`overlays.profile must contain exactly ${PROFILE_NAMES.join(", ")}`);
    }
    for (const profile of PROFILE_NAMES) {
      errors.push(...validateLabelArray(profiles[profile], `overlays.profile.${profile}`, seen));
    }
  }

  const domains = overlays?.domain;
  if (!domains || typeof domains !== "object" || Array.isArray(domains)) {
    errors.push("overlays.domain must be an object");
  } else {
    for (const name of Object.keys(domains).sort()) {
      if (!/^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(name)) {
        errors.push(`overlays.domain key is not kebab-case: ${name}`);
      }
      errors.push(...validateLabelArray(domains[name], `overlays.domain.${name}`, seen));
    }
  }

  if (errors.length > 0) throw new Error(`Invalid label manifest:\n- ${errors.join("\n- ")}`);
  return structuredClone(document);
}

export function resolveDesired(document, { profile = null, domains = [] } = {}) {
  const manifest = validateManifest(document);
  const labels = [...manifest.core];
  if (profile !== null) {
    if (!PROFILE_NAMES.includes(profile)) throw new Error(`Unknown profile overlay: ${profile}`);
    labels.push(...manifest.overlays.profile[profile]);
  }
  for (const domain of domains) {
    const overlay = manifest.overlays.domain[domain];
    if (!overlay) throw new Error(`Unknown domain overlay: ${domain}`);
    labels.push(...overlay);
  }
  labels.sort((left, right) => compareNames(left.name, right.name));
  return labels;
}

export function resolvePolicySelection(policy, manifest) {
  if (!policy || typeof policy !== "object" || Array.isArray(policy)) {
    throw new Error("Repository policy must be a JSON object");
  }
  if (!["template", "generated"].includes(policy.mode)) {
    throw new Error("Repository policy mode must be template or generated");
  }
  if (
    !Array.isArray(policy.label_domains)
    || policy.label_domains.some(domain => typeof domain !== "string")
  ) {
    throw new Error("Repository policy label_domains must be an array of strings");
  }
  const domains = [...policy.label_domains];
  if (new Set(domains).size !== domains.length) {
    throw new Error("Repository policy label_domains must not contain duplicates");
  }
  if (JSON.stringify(domains) !== JSON.stringify([...domains].sort(compareNames))) {
    throw new Error("Repository policy label_domains must be sorted case-insensitively");
  }
  for (const domain of domains) {
    if (!/^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(domain)) {
      throw new Error(`Repository policy label domain is not kebab-case: ${domain}`);
    }
    if (!manifest.overlays.domain[domain]) {
      throw new Error(`Repository policy selects unknown domain overlay: ${domain}`);
    }
  }
  if (policy.mode === "template") {
    if (policy.profile !== null || domains.length > 0) {
      throw new Error("Template policy must select no profile or domain overlays");
    }
    return { profile: null, domains: [] };
  }
  if (!PROFILE_NAMES.includes(policy.profile)) {
    throw new Error("Generated policy must select one supported profile overlay");
  }
  return { profile: policy.profile, domains };
}

export function planChanges(desiredLabels, liveLabels, { prune = true } = {}) {
  const desired = new Map(desiredLabels.map(label => [label.name.toLowerCase(), label]));
  const live = new Map(liveLabels.map(label => {
    const normalized = normalizedLiveLabel(label);
    return [normalized.name.toLowerCase(), normalized];
  }));
  const changes = [];

  for (const [key, target] of desired) {
    const current = live.get(key);
    if (!current) {
      changes.push({ action: "create", name: target.name, target });
      continue;
    }
    const fields = [];
    if (current.name !== target.name) fields.push("name");
    if (current.color !== target.color) fields.push("color");
    if (current.description !== target.description) fields.push("description");
    if (fields.length > 0) {
      changes.push({ action: "update", name: target.name, currentName: current.name, fields, target });
    }
  }

  if (prune) {
    for (const [key, current] of live) {
      if (!desired.has(key)) changes.push({ action: "delete", name: current.name, currentName: current.name });
    }
  }

  const order = { update: 0, create: 1, delete: 2 };
  changes.sort((left, right) => order[left.action] - order[right.action] || compareNames(left.name, right.name));
  return changes;
}

function parseArguments(argv) {
  const parsed = {
    manifest: DEFAULT_MANIFEST,
    policy: DEFAULT_POLICY,
    mode: null,
    live: null,
    repository: process.env.GITHUB_REPOSITORY || null,
    dryRun: false,
    selfTest: false
  };
  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index];
    if (argument === "--dry-run") {
      parsed.dryRun = true;
    } else if (argument === "--self-test") {
      parsed.selfTest = true;
    } else if (["--manifest", "--policy", "--mode", "--live", "--repository"].includes(argument)) {
      const value = argv[index + 1];
      if (!value || value.startsWith("--")) throw new Error(`${argument} requires a value`);
      parsed[argument.slice(2)] = value;
      index += 1;
    } else {
      throw new Error(`Unknown argument: ${argument}`);
    }
  }
  return parsed;
}

async function loadJson(path) {
  let text;
  try {
    text = await readFile(path, "utf8");
  } catch (error) {
    throw new Error(`Cannot read ${path}: ${error.message}`);
  }
  try {
    return JSON.parse(text);
  } catch (error) {
    throw new Error(`Invalid JSON in ${path}: ${error.message}`);
  }
}

function splitRepository(repository) {
  const parts = String(repository ?? "").split("/");
  if (parts.length !== 2 || parts.some(part => part.length === 0)) {
    throw new Error("repository must be in owner/name form");
  }
  return parts.map(encodeURIComponent);
}

// GitHub Actions provides the API address of the host it runs on.
function apiBase(env = process.env) {
  return String(env.GITHUB_API_URL ?? "").trim().replace(/\/+$/, "") || DEFAULT_API_URL;
}

function apiUrl(repository, suffix, env = process.env) {
  const [owner, repo] = splitRepository(repository);
  return `${apiBase(env)}/repos/${owner}/${repo}${suffix}`;
}

function sleep(milliseconds) {
  return new Promise(resolve => setTimeout(resolve, milliseconds));
}

function isRateLimited(response) {
  return response.status === 429
    || (
      response.status === 403
      && (
        response.headers.get("x-ratelimit-remaining") === "0"
        || response.headers.has("retry-after")
      )
    );
}

function retryDelayMs(response, attempt) {
  const header = response.headers.get("retry-after");
  const retryAfter = header !== null && header.trim() !== "" ? Number(header) : NaN;
  if (Number.isFinite(retryAfter) && retryAfter >= 0) {
    return Math.min(retryAfter * 1000, MAX_RETRY_DELAY_MS);
  }
  return Math.min(1000 * (2 ** attempt), MAX_RETRY_DELAY_MS);
}

async function githubRequest(repository, token, method, suffix, body) {
  for (let attempt = 0; ; attempt += 1) {
    const response = await fetch(apiUrl(repository, suffix), {
      method,
      headers: {
        Accept: "application/vnd.github+json",
        Authorization: `Bearer ${token}`,
        "X-GitHub-Api-Version": API_VERSION,
        "User-Agent": "canonical-vba-labels-sync"
      },
      body: body === undefined ? undefined : JSON.stringify(body)
    });
    if (response.ok) {
      if (response.status === 204) return null;
      return response.json();
    }

    const detail = await response.text();
    const retryable = method === "GET" && isRateLimited(response);
    if (!retryable || attempt >= MAX_RATE_LIMIT_RETRIES) {
      throw new Error(`${method} ${suffix} failed with HTTP ${response.status}: ${detail.slice(0, 500)}`);
    }
    await sleep(retryDelayMs(response, attempt));
  }
}

async function listLiveLabels(repository, token) {
  const labels = [];
  for (let page = 1; page <= MAX_LABEL_PAGES; page += 1) {
    const batch = await githubRequest(
      repository,
      token,
      "GET",
      `/labels?per_page=${LABELS_PER_PAGE}&page=${page}`
    );
    if (!Array.isArray(batch)) throw new Error("GitHub labels response is not an array");
    labels.push(...batch.map(normalizedLiveLabel));
    if (batch.length < LABELS_PER_PAGE) return labels;
  }
  throw new Error(`GitHub label pagination exceeded ${MAX_LABEL_PAGES} pages`);
}

function itemLabelNames(item) {
  if (item === null || typeof item !== "object" || !Array.isArray(item.labels)) {
    throw new Error("GitHub issues response has an item without a labels array");
  }
  return item.labels.map(label => {
    if (label === null || typeof label !== "object" || typeof label.name !== "string") {
      throw new Error("GitHub issues response has a malformed label");
    }
    return label.name.toLowerCase();
  });
}

function itemNumber(item) {
  if (!Number.isInteger(item?.number) || item.number < 1) {
    throw new Error("GitHub issues response has an item without a number");
  }
  return item.number;
}

// Returns, for each name among `names` that labels at least one issue or pull
// request, open or closed, `{ name, items, more }`: up to USAGE_SAMPLE item
// numbers, ascending, and whether further items use it. `request(suffix)`
// performs one GET. The labels filter cannot express a name that contains a
// comma, so those names are compared against every item, read page by page.
async function findLabelUsage(names, request) {
  const usage = [];
  const record = (name, numbers) => {
    const sorted = [...numbers].sort((left, right) => left - right);
    usage.push({ name, items: sorted.slice(0, USAGE_SAMPLE), more: sorted.length > USAGE_SAMPLE });
  };
  for (const name of names.filter(candidate => !candidate.includes(","))) {
    const batch = await request(
      `/issues?state=all&per_page=${USAGE_SAMPLE + 1}&labels=${encodeURIComponent(name)}`
    );
    if (!Array.isArray(batch)) throw new Error("GitHub issues response is not an array");
    if (batch.length > 0) record(name, batch.map(itemNumber));
  }
  const wanted = new Map(names.filter(name => name.includes(",")).map(name => [name.toLowerCase(), name]));
  if (wanted.size > 0) {
    const found = new Map();
    let complete = false;
    for (let page = 1; page <= MAX_ISSUE_PAGES && !complete; page += 1) {
      const batch = await request(`/issues?state=all&per_page=${ISSUES_PER_PAGE}&page=${page}`);
      if (!Array.isArray(batch)) throw new Error("GitHub issues response is not an array");
      for (const item of batch) {
        for (const name of itemLabelNames(item)) {
          if (wanted.has(name)) found.set(name, [...(found.get(name) ?? []), itemNumber(item)]);
        }
      }
      complete = batch.length < ISSUES_PER_PAGE;
    }
    if (!complete) throw new Error(`GitHub issue pagination exceeded ${MAX_ISSUE_PAGES} pages`);
    for (const [key, name] of wanted) if (found.has(key)) record(name, found.get(key));
  }
  return usage.sort((left, right) => compareNames(left.name, right.name));
}

// Reads, for each planned deletion, whether the label is still in use.
async function collectDeletionUsage(changes, request) {
  const deletions = changes.filter(change => change.action === "delete").map(change => change.currentName);
  if (deletions.length === 0) return [];
  return findLabelUsage(deletions, request);
}

// Pruning a label removes it from every issue and pull request and cannot be
// undone, so a planned deletion of a label still in use stops the run before
// any change is made.
function refuseUsedDeletions(usage) {
  if (usage.length > 0) {
    throw new Error(
      `Refusing to delete ${usage.length} label(s) still used by issues or pull requests: `
      + `${usage.map(entry => JSON.stringify(entry.name)).join(", ")}. Nothing was changed. `
      + "Remove the labels from those items, or add them to .github/labels.json."
    );
  }
}

async function applyChanges(repository, token, changes) {
  for (const change of changes) {
    if (change.action === "update") {
      await githubRequest(repository, token, "PATCH", `/labels/${encodeURIComponent(change.currentName)}`, {
        new_name: change.target.name,
        color: change.target.color,
        description: change.target.description
      });
    } else if (change.action === "create") {
      await githubRequest(repository, token, "POST", "/labels", change.target);
    } else if (change.action === "delete") {
      await githubRequest(repository, token, "DELETE", `/labels/${encodeURIComponent(change.currentName)}`);
    }
  }
}

function renderLabelNames(labels) {
  if (labels.length === 0) return "none";
  return labels.map(label => inlineCode(label.name)).join(", ");
}

// Where a label planned for deletion is used, e.g. "#3, #7 and others".
function renderUsage(entry) {
  const numbers = entry.items.map(number => `#${number}`).join(", ");
  return `${numbers}${entry.more ? " and others" : ""}`;
}

function renderSummary({
  mode,
  manifest,
  manifestPath,
  policyPath,
  desiredLabels,
  profile,
  domains,
  changes = [],
  verified = null,
  usage = [],
  dryRun = false,
  error = null
}) {
  const coreLabels = manifest?.core ?? [];
  const profileLabels = profile === null ? [] : manifest?.overlays?.profile?.[profile] ?? [];
  const domainLabels = domains.flatMap(domain => manifest?.overlays?.domain?.[domain] ?? []);
  const resolvedLabels = desiredLabels ?? [];
  const counts = Object.fromEntries(["create", "update", "delete"].map(action => [
    action,
    changes.filter(change => change.action === action).length
  ]));
  const lines = [
    "# Repository label synchronization",
    "",
    `- Mode: ${inlineCode(mode)}${dryRun ? " (dry run)" : ""}`,
    `- Policy source: ${inlineCode(policyPath ?? "unavailable")}`,
    `- Label manifest: ${inlineCode(manifestPath ?? "unavailable")}`,
    `- Schema: **${manifest?.schema_version ?? "unavailable"}**`,
    `- Core / selected labels: **${coreLabels.length} / ${resolvedLabels.length}**`,
    `- Profile overlay: **${profile ?? "none"}**`,
    `- Domain overlays: **${domains.length > 0 ? domains.join(", ") : "none"}**`,
    `- Prune: **${manifest?.prune ? "enabled" : "disabled"}**`,
    `- Planned changes: **${changes.length}**`,
    `- Create / update / delete: **${counts.create} / ${counts.update} / ${counts.delete}**`
  ];
  const blocked = new Map(usage.map(entry => [entry.name.toLowerCase(), entry]));
  if (blocked.size > 0) lines.push(`- Deletions blocked by labels in use: **${blocked.size}**`);
  if (verified !== null) lines.push(`- Post-run exact match: **${verified ? "yes" : "no"}**`);
  lines.push(
    "",
    "## Resolved label catalog",
    "",
    `- Core labels (${coreLabels.length}): ${renderLabelNames(coreLabels)}`,
    `- Profile labels (${profile ?? "none"}; ${profileLabels.length}): ${renderLabelNames(profileLabels)}`,
    `- Domain labels (${domains.length > 0 ? domains.join(", ") : "none"}; ${domainLabels.length}): ${renderLabelNames(domainLabels)}`,
    `- Complete resolved set (${resolvedLabels.length}): ${renderLabelNames(resolvedLabels)}`
  );
  if (error) {
    lines.push("", "## Failure", "", inlineCode(error));
  } else if (changes.length === 0) {
    lines.push("", "No label changes are required.");
  } else {
    lines.push("", "| Action | Label | Detail |", "|---|---|---|");
    for (const change of changes) {
      let detail = change.action === "update" ? change.fields.join(", ") : change.action === "delete" ? "outside selected manifest" : "missing live label";
      const entry = change.action === "delete" ? blocked.get(change.currentName.toLowerCase()) : undefined;
      if (entry) detail += `; blocked, used by ${renderUsage(entry)}`;
      lines.push(`| ${markdownEscape(change.action)} | ${inlineCode(change.name, { table: true })} | ${markdownEscape(detail)} |`);
    }
    if (blocked.size > 0) {
      lines.push(
        "",
        `A reconcile run would refuse ${blocked.size} deletion(s) of labels still used by issues or pull requests `
        + "and change nothing. Remove the labels from those items, or add them to .github/labels.json."
      );
    }
  }
  return `${lines.join("\n")}\n`;
}

async function publishSummary(summary) {
  process.stdout.write(summary);
  if (process.env.GITHUB_STEP_SUMMARY) await appendFile(process.env.GITHUB_STEP_SUMMARY, summary, "utf8");
}

function expectFailure(operation, pattern) {
  assert.throws(operation, pattern);
}

async function runSelfTest(manifestPath, policyPath) {
  for (const [header, attempt, expected] of [
    [null, 0, 1000], [null, 2, 4000], ["", 1, 2000],
    ["invalid", 1, 2000], ["-1", 1, 2000], ["0", 2, 0],
    ["2", 0, 2000], ["999999", 0, MAX_RETRY_DELAY_MS],
    [null, 20, MAX_RETRY_DELAY_MS]
  ]) {
    const headers = new Headers(header === null ? {} : { "retry-after": header });
    assert.equal(retryDelayMs({ headers }, attempt), expected, `Retry-After: ${header}`);
  }

  const baseline = await loadJson(manifestPath);
  const manifest = validateManifest(baseline);
  const policy = await loadJson(policyPath);
  const desired = resolveDesired(manifest);

  const currentSelection = resolvePolicySelection(policy, manifest);
  assert.deepEqual(
    currentSelection,
    policy.mode === "template"
      ? { profile: null, domains: [] }
      : { profile: policy.profile, domains: policy.label_domains }
  );
  const templatePolicy = structuredClone(policy);
  templatePolicy.mode = "template";
  templatePolicy.profile = null;
  templatePolicy.label_domains = [];
  assert.deepEqual(resolvePolicySelection(templatePolicy, manifest), { profile: null, domains: [] });

  const generatedPolicy = structuredClone(templatePolicy);
  generatedPolicy.mode = "generated";
  generatedPolicy.profile = "library";
  assert.deepEqual(resolvePolicySelection(generatedPolicy, manifest), { profile: "library", domains: [] });

  const wrongPolicyMode = structuredClone(policy);
  wrongPolicyMode.mode = "unknown";
  expectFailure(() => resolvePolicySelection(wrongPolicyMode, manifest), /mode/);

  const selectedTemplateProfile = structuredClone(templatePolicy);
  selectedTemplateProfile.profile = "library";
  expectFailure(() => resolvePolicySelection(selectedTemplateProfile, manifest), /Template policy/);

  const missingGeneratedProfile = structuredClone(generatedPolicy);
  missingGeneratedProfile.profile = null;
  expectFailure(() => resolvePolicySelection(missingGeneratedProfile, manifest), /supported profile/);

  const invalidPolicyDomains = structuredClone(generatedPolicy);
  invalidPolicyDomains.label_domains = ["Unknown Domain"];
  expectFailure(() => resolvePolicySelection(invalidPolicyDomains, manifest), /kebab-case/);

  const missingRoot = structuredClone(baseline);
  delete missingRoot.prune;
  expectFailure(() => validateManifest(missingRoot), /root must contain exactly/);

  const wrongSchema = structuredClone(baseline);
  wrongSchema.schema_version = 2;
  expectFailure(() => validateManifest(wrongSchema), /schema_version/);

  const invalidPrune = structuredClone(baseline);
  invalidPrune.prune = "true";
  expectFailure(() => validateManifest(invalidPrune), /prune must be a boolean/);

  const missingField = structuredClone(baseline);
  delete missingField.core[0].description;
  expectFailure(() => validateManifest(missingField), /name, color, and description/);

  const invalidColor = structuredClone(baseline);
  invalidColor.core[0].color = "#c24e00";
  expectFailure(() => validateManifest(invalidColor), /uppercase hexadecimal/);

  const duplicateName = structuredClone(baseline);
  duplicateName.core[1].name = baseline.core[0].name.toUpperCase();
  expectFailure(() => validateManifest(duplicateName), /duplicates/);

  const overlayDuplicate = structuredClone(baseline);
  overlayDuplicate.overlays.profile.library.push(structuredClone(baseline.core[0]));
  expectFailure(() => validateManifest(overlayDuplicate), /duplicates/);

  const longDescription = structuredClone(baseline);
  longDescription.core[0].description = "x".repeat(101);
  expectFailure(() => validateManifest(longDescription), /100-character/);

  const unsorted = structuredClone(baseline);
  [unsorted.core[0], unsorted.core[1]] = [unsorted.core[1], unsorted.core[0]];
  expectFailure(() => validateManifest(unsorted), /sorted/);

  expectFailure(() => resolveDesired(manifest, { profile: "unknown" }), /Unknown profile/);
  expectFailure(() => resolveDesired(manifest, { domains: ["unknown"] }), /Unknown domain/);
  assert.equal(planChanges(desired, desired, { prune: manifest.prune }).length, 0, "idempotent plan must be empty");

  assert.deepEqual(planChanges(desired, desired.slice(1)).map(change => change.action), ["create"]);

  const changed = structuredClone(desired);
  changed[0].color = "000000";
  changed[0].description = "changed";
  const changedPlan = planChanges(desired, changed);
  assert.deepEqual(changedPlan.map(change => change.action), ["update"]);
  assert.deepEqual(changedPlan[0].fields, ["color", "description"]);

  const extra = [...structuredClone(desired), { name: "extra", color: "ABCDEF", description: "extra" }];
  assert.deepEqual(planChanges(desired, extra, { prune: true }).map(change => change.action), ["delete"]);
  assert.equal(planChanges(desired, extra, { prune: false }).length, 0);

  const caseDrift = structuredClone(desired);
  caseDrift[0].name = desired[0].name.toUpperCase();
  assert.deepEqual(planChanges(desired, caseDrift)[0].fields, ["name"]);

  const overlay = structuredClone(baseline);
  overlay.overlays.profile.library = [{ name: "library", color: "654321", description: "Library profile label" }];
  overlay.overlays.domain.alpha = [{ name: "alpha", color: "123ABC", description: "Alpha domain label" }];
  overlay.overlays.domain.example = [{ name: "example", color: "ABCDEF", description: "Example domain label" }];
  const selectedOverlay = validateManifest(overlay);
  const selectedLabels = resolveDesired(selectedOverlay, { profile: "library", domains: ["example"] });
  assert.equal(selectedLabels.length, desired.length + 2);

  const domainPolicy = structuredClone(generatedPolicy);
  domainPolicy.label_domains = ["example"];
  assert.deepEqual(resolvePolicySelection(domainPolicy, overlay), { profile: "library", domains: ["example"] });

  const unknownDomainPolicy = structuredClone(generatedPolicy);
  unknownDomainPolicy.label_domains = ["unknown"];
  expectFailure(() => resolvePolicySelection(unknownDomainPolicy, manifest), /unknown domain/);

  const duplicateDomainPolicy = structuredClone(generatedPolicy);
  duplicateDomainPolicy.label_domains = ["example", "example"];
  expectFailure(() => resolvePolicySelection(duplicateDomainPolicy, overlay), /duplicates/);

  const unsortedDomainPolicy = structuredClone(generatedPolicy);
  unsortedDomainPolicy.label_domains = ["example", "alpha"];
  expectFailure(() => resolvePolicySelection(unsortedDomainPolicy, overlay), /sorted/);

  const summary = renderSummary({
    mode: "plan",
    manifest: selectedOverlay,
    manifestPath: "fixture-labels.json",
    policyPath: "fixture-policy.json",
    desiredLabels: selectedLabels,
    profile: "library",
    domains: ["example"],
    dryRun: true
  });
  assert.match(summary, /Policy source: `fixture-policy\.json`/);
  assert.match(summary, /Label manifest: `fixture-labels\.json`/);
  assert.match(summary, /Profile labels \(library; 1\): `library`/);
  assert.match(summary, /Domain labels \(example; 1\): `example`/);
  assert.match(summary, new RegExp(`Complete resolved set \\(${selectedLabels.length}\\):`));
  for (const label of selectedLabels) assert.ok(summary.includes(inlineCode(label.name)), label.name);

  selfTestMarkdown();
  selfTestSummaryMarkdown();

  await selfTestDeletionGuard();

  for (const [value, expected] of [
    [undefined, DEFAULT_API_URL], ["", DEFAULT_API_URL], ["  ", DEFAULT_API_URL],
    ["https://github.example.com/api/v3", "https://github.example.com/api/v3"],
    ["https://github.example.com/api/v3/", "https://github.example.com/api/v3"],
    ["http://127.0.0.1:8080//", "http://127.0.0.1:8080"]
  ]) {
    const env = value === undefined ? {} : { GITHUB_API_URL: value };
    assert.equal(apiUrl("owner/repo", "/labels", env), `${expected}/repos/owner/repo/labels`, `GITHUB_API_URL: ${value}`);
  }

  await selfTestReconcileEndToEnd(manifestPath, policyPath, resolveDesired(manifest, currentSelection));

  process.stdout.write("SELF-TEST PASS: retry delays, validation, policy, overlay, reconciliation, Markdown summary, deletion guard, API URL and end-to-end reconcile and dry-run fixtures\n");
}

// Label names, paths and error text with backticks, pipes and line breaks
// stay inside their code span, their table cell and their line.
function selfTestSummaryMarkdown() {
  const names = ["wip `x`", "a ``b`` c", "`edge", "edge`", "a|b", "cr\rlf\ncrlf\r\nend"];
  const labels = names.map(name => ({ name, color: "ABCDEF", description: "d" }));
  const summary = renderSummary({
    mode: "plan",
    manifest: null,
    manifestPath: "dir`s/labels|v1.json",
    policyPath: "policy\r\n.json",
    desiredLabels: labels,
    profile: null,
    domains: [],
    changes: [
      ...labels.map(label => ({ action: "create", name: label.name, target: label })),
      { action: "update", name: "a|b", currentName: "A|B", fields: ["name", "color"], target: labels[4] }
    ]
  });
  const lines = summary.split("\n");
  assert.ok(lines.includes("- Label manifest: ``dir`s/labels|v1.json``"));
  assert.ok(lines.includes("- Policy source: `policy .json`"));
  assert.ok(lines.includes(
    "- Complete resolved set (6): `` wip `x` ``, ```a ``b`` c```, `` `edge ``, `` edge` ``, `a|b`, `cr lf crlf end`"
  ));
  const rows = lines.filter(line => line.startsWith("| create ") || line.startsWith("| update "));
  assert.deepEqual(rows, [
    "| create | `` wip `x` `` | missing live label |",
    "| create | ```a ``b`` c``` | missing live label |",
    "| create | `` `edge `` | missing live label |",
    "| create | `` edge` `` | missing live label |",
    "| create | `a\\|b` | missing live label |",
    "| create | `cr lf crlf end` | missing live label |",
    "| update | `a\\|b` | name, color |"
  ]);

  const failure = renderSummary({
    mode: "failed",
    manifest: null,
    manifestPath: null,
    policyPath: null,
    desiredLabels: [],
    profile: null,
    domains: [],
    error: "Invalid label manifest:\r\n- core[0].name `x``y` | bad\r- next\nend`"
  });
  const blockedSummary = renderSummary({
    mode: "reconcile",
    manifest: null,
    manifestPath: null,
    policyPath: null,
    desiredLabels: [],
    profile: null,
    domains: [],
    dryRun: true,
    changes: [
      { action: "delete", name: "old|`x`", currentName: "old|`x`" },
      { action: "delete", name: "spare", currentName: "spare" }
    ],
    usage: [{ name: "OLD|`X`", items: [3, 7], more: true }]
  }).split("\n");
  assert.ok(blockedSummary.includes("- Deletions blocked by labels in use: **1**"));
  assert.deepEqual(blockedSummary.filter(line => line.startsWith("| delete ")), [
    "| delete | `` old\\|`x` `` | outside selected manifest; blocked, used by #3, #7 and others |",
    "| delete | `spare` | outside selected manifest |"
  ]);
  assert.ok(blockedSummary.some(line => line.startsWith("A reconcile run would refuse 1 deletion(s)")));

  assert.ok(failure.endsWith("## Failure\n\n``` Invalid label manifest: - core[0].name `x``y` | bad - next end` ```\n"), failure);
}

// A fake GET for the deletion guard: `items` are the repository's issues and
// pull requests; the labels filter and pagination behave as GitHub's do.
function fakeIssues(items, calls = []) {
  return async suffix => {
    calls.push(suffix);
    const url = new URL(suffix, "https://api.invalid");
    assert.equal(url.pathname, "/issues");
    assert.equal(url.searchParams.get("state"), "all");
    const filter = url.searchParams.get("labels");
    const perPage = Number(url.searchParams.get("per_page"));
    const page = Number(url.searchParams.get("page") ?? "1");
    const matching = filter === null
      ? items
      : items.filter(item => item.labels.some(label => label.name.toLowerCase() === filter.toLowerCase()));
    return matching.slice((page - 1) * perPage, page * perPage);
  };
}

async function selfTestDeletionGuard() {
  let nextNumber = 1;
  const labelled = (...names) => ({ number: nextNumber++, labels: names.map(name => ({ name })) });
  const remove = (...names) => names.map(name => ({ action: "delete", name, currentName: name }));
  const guard = async (changes, request) => refuseUsedDeletions(await collectDeletionUsage(changes, request));

  const noCalls = [];
  assert.deepEqual(await collectDeletionUsage([{ action: "create", name: "new" }], fakeIssues([], noCalls)), []);
  assert.deepEqual(noCalls, [], "no deletion, no request");

  await guard(remove("unused"), fakeIssues([labelled("bug")]));

  const usedItems = [labelled("Legacy"), labelled("bug", "old"), labelled("old")];
  assert.deepEqual(await collectDeletionUsage(remove("old", "legacy", "unused"), fakeIssues(usedItems)), [
    { name: "legacy", items: [usedItems[0].number], more: false },
    { name: "old", items: [usedItems[1].number, usedItems[2].number], more: false }
  ]);
  await assert.rejects(
    guard(remove("legacy", "old", "unused"), fakeIssues(usedItems)),
    error => error.message === 'Refusing to delete 2 label(s) still used by issues or pull requests: '
      + '"legacy", "old". Nothing was changed. Remove the labels from those items, or add them to .github/labels.json.'
  );

  const many = Array.from({ length: 150 }, () => labelled("bug"));
  many.push(labelled("Area, Legacy"));
  const calls = [];
  assert.deepEqual(
    await findLabelUsage(["area, legacy", "a, b"], fakeIssues(many, calls)),
    [{ name: "area, legacy", items: [many.at(-1).number], more: false }]
  );
  assert.deepEqual(calls, [
    "/issues?state=all&per_page=100&page=1",
    "/issues?state=all&per_page=100&page=2"
  ], "comma names read every page once, unfiltered");

  const plainCalls = [];
  await findLabelUsage(["needs triage"], fakeIssues([], plainCalls));
  assert.deepEqual(plainCalls, [`/issues?state=all&per_page=${USAGE_SAMPLE + 1}&labels=needs%20triage`]);

  const crowded = Array.from({ length: USAGE_SAMPLE + 2 }, () => labelled("busy", "x, y"));
  const sample = crowded.slice(0, USAGE_SAMPLE).map(item => item.number);
  assert.deepEqual(await findLabelUsage(["busy", "x, y"], fakeIssues(crowded)), [
    { name: "busy", items: sample, more: true },
    { name: "x, y", items: sample, more: true }
  ]);
  const exact = Array.from({ length: USAGE_SAMPLE }, () => labelled("full"));
  assert.equal((await findLabelUsage(["full"], fakeIssues(exact)))[0].more, false);

  await assert.rejects(findLabelUsage(["old"], async () => ({ message: "x" })), /not an array/);
  await assert.rejects(findLabelUsage(["old"], async () => [{ labels: [] }]), /without a number/);
  await assert.rejects(findLabelUsage(["a, b"], async () => ({})), /not an array/);
  await assert.rejects(findLabelUsage(["a, b"], async () => [{ number: 1 }]), /labels array/);
  await assert.rejects(findLabelUsage(["a, b"], async () => [{ labels: [{ id: 1 }] }]), /malformed label/);
  await assert.rejects(findLabelUsage(["a, b"], async () => [{ labels: ["a, b"] }]), /malformed label/);
  await assert.rejects(findLabelUsage(["a, b"], async () => [{ number: "1", labels: [{ name: "A, B" }] }]), /without a number/);
  const fullPages = [];
  await assert.rejects(
    findLabelUsage(["a, b"], async suffix => {
      fullPages.push(suffix);
      return Array.from({ length: ISSUES_PER_PAGE }, () => labelled("bug"));
    }),
    /pagination exceeded 100 pages/
  );
  assert.equal(fullPages.length, MAX_ISSUE_PAGES);
}

// Runs this script in reconcile mode against a local API that holds the
// selected labels plus "legacy": the run deletes "legacy" only when no issue
// or pull request uses it, and a dry run reports the blocked deletion without
// a change. The fixture uses its own copy of the manifest with
// pruning enabled, so it holds whatever the repository's prune setting is.
async function selfTestReconcileEndToEnd(manifestPath, policyPath, desired) {
  const folder = await mkdtemp(join(tmpdir(), "labels-sync-"));
  try {
    const pruning = { ...(await loadJson(manifestPath)), prune: true };
    const fixtureManifest = join(folder, "labels.json");
    await writeFile(fixtureManifest, JSON.stringify(pruning), "utf8");
    await reconcileAgainstLocalApi(fixtureManifest, policyPath, desired);
  } finally {
    await rm(folder, { recursive: true, force: true });
  }
}

async function reconcileAgainstLocalApi(manifestPath, policyPath, desired) {
  for (const [used, dryRun] of [[true, false], [false, false], [true, true]]) {
    let live = [...structuredClone(desired), { name: "legacy", color: "ABCDEF", description: "Old label" }];
    const requests = [];
    const server = createServer((request, response) => {
      requests.push(`${request.method} ${request.url}`);
      const url = new URL(request.url, "http://localhost");
      response.setHeader("content-type", "application/json");
      if (request.method === "GET" && url.pathname === "/repos/owner/repo/labels") {
        response.end(JSON.stringify(url.searchParams.get("page") === "1" ? live : []));
      } else if (request.method === "GET" && url.pathname === "/repos/owner/repo/issues") {
        response.end(JSON.stringify(used ? [{ number: 7, labels: [{ name: "Legacy" }] }] : []));
      } else if (request.method === "DELETE" && url.pathname === "/repos/owner/repo/labels/legacy") {
        live = live.filter(label => label.name !== "legacy");
        response.statusCode = 204;
        response.end();
      } else {
        response.statusCode = 404;
        response.end("{}");
      }
    });
    await new Promise(resolve => server.listen(0, "127.0.0.1", resolve));
    const env = { ...process.env, GITHUB_API_URL: `http://127.0.0.1:${server.address().port}/`, GITHUB_TOKEN: "test" };
    delete env.GITHUB_STEP_SUMMARY;
    const result = await new Promise(resolve => execFile(
      process.execPath,
      [fileURLToPath(import.meta.url), "--manifest", manifestPath, "--policy", policyPath,
        "--mode", "reconcile", "--repository", "owner/repo", ...(dryRun ? ["--dry-run"] : [])],
      { env },
      (error, stdout, stderr) => resolve({ code: error ? error.code : 0, stdout, stderr })
    ));
    await new Promise(resolve => server.close(resolve));
    const deletes = requests.filter(line => line.startsWith("DELETE "));
    const writes = requests.filter(line => !line.startsWith("GET "));
    assert.ok(
      requests.includes(`GET /repos/owner/repo/issues?state=all&per_page=${USAGE_SAMPLE + 1}&labels=legacy`),
      requests.join("\n")
    );
    if (dryRun) {
      assert.equal(result.code, 0, result.stderr);
      assert.deepEqual(writes, [], "a dry run changes nothing");
      assert.match(result.stdout, /Mode: `reconcile` \(dry run\)/);
      assert.match(result.stdout, /Deletions blocked by labels in use: \*\*1\*\*/);
      assert.ok(
        result.stdout.includes("| delete | `legacy` | outside selected manifest; blocked, used by #7 |"),
        result.stdout
      );
      assert.match(result.stdout, /A reconcile run would refuse 1 deletion\(s\)/);
      assert.doesNotMatch(result.stdout, /Post-run exact match/);
    } else if (used) {
      assert.equal(result.code, 1, result.stdout);
      assert.equal(
        result.stderr,
        'Refusing to delete 1 label(s) still used by issues or pull requests: "legacy". Nothing was changed. '
        + "Remove the labels from those items, or add them to .github/labels.json.\n"
      );
      assert.deepEqual(writes, [], "a label in use is never deleted, and nothing else changes");
    } else {
      assert.equal(result.code, 0, result.stderr);
      assert.deepEqual(deletes, ["DELETE /repos/owner/repo/labels/legacy"]);
      assert.match(result.stdout, /Post-run exact match: \*\*yes\*\*/);
    }
  }
}

async function main() {
  const options = parseArguments(process.argv.slice(2));
  if (options.selfTest) {
    await runSelfTest(options.manifest, options.policy);
    return;
  }

  const manifest = validateManifest(await loadJson(options.manifest));
  const selection = resolvePolicySelection(await loadJson(options.policy), manifest);
  const desired = resolveDesired(manifest, selection);
  const summaryBase = {
    manifest,
    manifestPath: options.manifest,
    policyPath: options.policy,
    desiredLabels: desired,
    profile: selection.profile,
    domains: selection.domains
  };

  if (options.mode === "validate") {
    await publishSummary(renderSummary({ mode: "validate", ...summaryBase }));
    return;
  }

  let live;
  if (options.mode === "plan") {
    if (!options.live) throw new Error("--mode plan requires --live");
    const liveDocument = await loadJson(options.live);
    live = Array.isArray(liveDocument) ? liveDocument : liveDocument.labels;
    if (!Array.isArray(live)) throw new Error("live fixture must be an array or contain a labels array");
  } else if (options.mode === "reconcile") {
    if (!options.repository) throw new Error("--mode reconcile requires --repository or GITHUB_REPOSITORY");
    if (!process.env.GITHUB_TOKEN) throw new Error("--mode reconcile requires GITHUB_TOKEN");
    live = await listLiveLabels(options.repository, process.env.GITHUB_TOKEN);
  } else {
    throw new Error("--mode must be validate, plan, or reconcile");
  }

  const changes = planChanges(desired, live, { prune: manifest.prune });
  let usage = [];
  if (options.mode === "reconcile") {
    usage = await collectDeletionUsage(
      changes,
      suffix => githubRequest(options.repository, process.env.GITHUB_TOKEN, "GET", suffix)
    );
    if (!options.dryRun) refuseUsedDeletions(usage);
  }
  if (options.mode === "plan" || options.dryRun) {
    await publishSummary(renderSummary({ mode: options.mode, ...summaryBase, changes, usage, dryRun: options.dryRun }));
    return;
  }

  await applyChanges(options.repository, process.env.GITHUB_TOKEN, changes);
  const after = await listLiveLabels(options.repository, process.env.GITHUB_TOKEN);
  const remaining = planChanges(desired, after, { prune: manifest.prune });
  const verified = remaining.length === 0;
  await publishSummary(renderSummary({ mode: "reconcile", ...summaryBase, changes, verified }));
  if (!verified) throw new Error(`Post-run verification found ${remaining.length} remaining difference(s)`);
}

main().catch(async error => {
  const message = error instanceof Error ? error.message : String(error);
  try {
    await publishSummary(renderSummary({
      mode: "failed",
      manifest: null,
      manifestPath: null,
      policyPath: null,
      desiredLabels: [],
      profile: null,
      domains: [],
      error: message
    }));
  } finally {
    process.stderr.write(`${message}\n`);
    process.exitCode = 1;
  }
});
