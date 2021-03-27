//
// Copyright (c) Microsoft.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
//

// This JS file is shared between ospo-witness-event & ospo-opensource-repos
// where reports are generated. It helps take a minimum viable consolidated
// report, containing issue type definitions and also the entities impacted
// by that issue, and a simple recipient list, and explodes that into a set
// of reports for a specific recipient.
//
// Any changes should be fully synchronized between the two repos. There is
// a "schema version" saved inside the metadata that is validated by the
// witness event digest processor before using a consolidated report. While
// bug fixes to this file can be made without changing the schema version,
// any changes to the consolidated format should take into account
// compatibility.

function addEntityToRecipientMap(
  map,
  providerName,
  recipient,
  entity,
  definitions,
  options
) {
  const filterDefinitionCategories = options.filterDefinitionCategories
    ? new Set(options.filterDefinitionCategories)
    : false;
  const simplifiedRecipientName = `${recipient.type}:${recipient.value}`;
  let recipientView = map.get(simplifiedRecipientName);
  if (!recipientView) {
    recipientView = {
      reasons: new Set(),
    };
    map.set(simplifiedRecipientName, recipientView);
  }
  const issueNames = Object.getOwnPropertyNames(entity.issues);
  for (let i = 0; i < issueNames.length; i++) {
    const issueName = issueNames[i];
    if (
      recipient.specific &&
      recipient.specific.issueNames &&
      !recipient.specific.issueNames.has(issueName)
    ) {
      continue;
    }
    const definition = definitions[issueName];
    if (
      filterDefinitionCategories &&
      definition.category &&
      !filterDefinitionCategories.has(definition.category)
    ) {
      continue;
    }
    if (recipient.reasons) {
      for (let i = 0; i < recipient.reasons.length; i++) {
        recipientView.reasons.add(recipient.reasons[i]);
      }
    }
    if (!recipientView[issueName]) {
      const entry: IEntry = {
        definition: definition,
      };
      if (definition.hasTable) {
        entry.table = {
          rows: [],
        };
        Object.assign(entry.table, definition.table);
      }
      if (definition.hasList) {
        entry.list = {
          listItems: [],
        };
        Object.assign(entry.list, definition.list);
      }
      recipientView[issueName] = entry;
    }
    let entry = recipientView[issueName];
    const entityIssue = entity.issues[issueName];
    const specificItems =
      entity.specific && entity.specific.issueItems
        ? entity.specific.issueItems
        : null;
    if (definition.hasTable) {
      fillFrom(entityIssue, 'rows', entry.table, entity, specificItems);
    }
    if (definition.hasList) {
      fillFrom(entityIssue, 'listItems', entry.list, entity, specificItems);
    }
  }
}

function fillFrom(object, property, target, entity, specificItems) {
  const source = object[property];
  if (
    source &&
    Array.isArray(source) &&
    Array.isArray(target[property]) &&
    source.length
  ) {
    const targetArray = target[property];
    for (let i = 0; i < source.length; i++) {
      const sourceItem = source[i];
      if (specificItems && !specificItems.has(sourceItem)) {
        continue;
      }
      let lineItem =
        typeof source[i] === 'object'
          ? Object.assign({}, sourceItem)
          : { text: sourceItem };
      if (!lineItem.entityName && entity.name) {
        lineItem.entityName = entity.name;
      }
      targetArray.push(lineItem);
    }
  }
}

function identifyAdditionalRecipients(entity, recipients) {
  const additionals = [];
  const additionalEntries = new Map();
  const issues = entity.issues;
  if (!issues) {
    return additionals;
  }
  const issueNames = Object.getOwnPropertyNames(issues);
  for (let i = 0; i < issueNames.length; i++) {
    const issueName = issueNames[i];
    const issue = entity.issues[issueNames[i]];
    let items = null;
    if (issue.listItems && issue.listItems.length) {
      items = issue.listItems;
    } else if (issue.rows && issue.rows.length) {
      items = issue.rows;
    }
    if (items) {
      for (let j = 0; j < items.length; j++) {
        const item = items[j];
        if (item.additionalRecipients) {
          for (let k = 0; k < item.additionalRecipients.length; k++) {
            const recipient = item.additionalRecipients[k];
            let found = null;
            for (let l = 0; l < recipients.length; l++) {
              const existing = recipients[l];
              if (
                existing.type === recipient.type &&
                existing.value === recipient.value
              ) {
                found = existing;
                break;
              }
            }
            if (found) {
              const reasonSet = new Set(found.reasons);
              for (let m = 0; m < recipient.reasons.length; m++) {
                reasonSet.add(recipient.reasons[m]);
              }
              found.reasons = Array.from(reasonSet.values());
            } else {
              const combined = `:${recipient.type}:${recipient.value}:`;
              let entry = additionalEntries.get(combined);
              if (!entry) {
                entry = Object.assign(
                  {
                    specific: {
                      issueNames: new Set(),
                      issueItems: new Set(),
                    },
                  },
                  item.additionalRecipients[k]
                );
                additionalEntries.set(combined, entry);
                additionals.push(entry);
              }
              entry.specific.issueNames.add(issueName);
              entry.specific.issueItems.add(item);
            }
          }
        }
      }
    }
  }
  return additionals;
}

function deduplicateRecipients(recipients) {
  const visited = new Map();
  const r = [];
  for (let i = 0; i < recipients.length; i++) {
    const recipient = recipients[i];
    const combined = `:${recipient.type}:${recipient.value}:`;
    let deduplicatedEntry = visited.get(combined);
    if (!deduplicatedEntry) {
      const clonedRecipient = Object.assign({}, recipient);
      delete clonedRecipient.reasons;
      r.push(clonedRecipient);
      deduplicatedEntry = {
        reasons: new Set(),
        clone: clonedRecipient,
      };
      visited.set(combined, deduplicatedEntry);
    }
    if (recipient.reasons) {
      for (let j = 0; j < recipient.reasons.length; j++) {
        deduplicatedEntry.reasons.add(recipient.reasons[j]);
      }
      deduplicatedEntry.clone.reasons = Array.from(
        deduplicatedEntry.reasons.values()
      );
    }
  }
  return r;
}

export function buildConsolidatedMap(consolidated, options?) {
  options = options || {};
  const byRecipient = new Map();
  const providerNames = Object.getOwnPropertyNames(consolidated);
  for (let i = 0; i < providerNames.length; i++) {
    const providerName = providerNames[i];
    const dataset = consolidated[providerName];
    if (typeof dataset !== 'object' || providerName === 'metadata') {
      continue;
    }
    const definitions = {};
    const providerByName = new Map();
    for (let x = 0; x < dataset.definitions.length; x++) {
      const d = dataset.definitions[x];
      definitions[d.name] = d;
    }
    if (dataset.entities && dataset.entities.length) {
      for (let j = 0; j < dataset.entities.length; j++) {
        const entity = dataset.entities[j];
        const recipients = deduplicateRecipients(
          entity && entity.recipients ? entity.recipients : []
        );
        const additionalRecipients = identifyAdditionalRecipients(
          entity,
          recipients
        );
        const allRecipients = recipients.concat(additionalRecipients);
        const entityClone = Object.assign({}, entity);
        delete entityClone.recipients;
        for (let k = 0; k < allRecipients.length; k++) {
          const recipient = allRecipients[k];
          addEntityToRecipientMap(
            providerByName,
            providerName,
            recipient,
            entityClone,
            definitions,
            options
          );
        }
      }
    }
    for (let recipient of providerByName.keys()) {
      const values = providerByName.get(recipient);
      if (!byRecipient.has(recipient)) {
        const recipientEntries = [];
        ((recipientEntries as any) as IReasonsSet).reasons = new Set();
        byRecipient.set(recipient, recipientEntries);
      }
      const entry = byRecipient.get(recipient);
      if (values.reasons && entry.reasons) {
        for (let reason of values.reasons) {
          entry.reasons.add(reason);
        }
      }
      for (let d = 0; d < dataset.definitions.length; d++) {
        const definition = dataset.definitions[d];
        if (values[definition.name]) {
          entry.push(values[definition.name]);
        }
      }
    }
  }
  // Reduce the set of reasons down to an array; remove empty reports
  const keys = Array.from(byRecipient.keys());
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    const value = byRecipient.get(key);
    if (value.length === 0) {
      byRecipient.delete(key);
      continue;
    }
    if (value.reasons && value.reasons.add) {
      value.reasons = Array.from(value.reasons.values());
    }
  }
  return byRecipient;
}

interface IReasonsSet {
  reasons: Set<any>;
}

interface IEntry {
  definition: any;
  table?: {
    rows: any[];
  };
  list?: {
    listItems: any[];
  };
}
