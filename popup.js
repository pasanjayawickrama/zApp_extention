const output = document.getElementById("output");

const PRIORITY_BUCKETS = ["1 - Critical", "2 - High", "3 - Moderate", "4 - Low"];

function safeJsonParse(value) {
  if (typeof value !== "string") return null;
  try {
    return JSON.parse(value);
  } catch {
    return null;
  }
}

function computePriorityCounts(graphqlResponseData) {
  const counts = {
    "1 - Critical": 0,
    "2 - High": 0,
    "3 - Moderate": 0,
    "4 - Low": 0,
    Unknown: 0
  };

  const results =
    graphqlResponseData?.data?.GlideRecord_Query?.ui_notification_inbox?._results || [];

  for (const item of results) {
    const payloadValue = item?.payload?.value;
    const payloadObj = safeJsonParse(payloadValue);

    const caseNumber = payloadObj?.displayValue;
    const priorityDisplay = payloadObj?.changes?.priority?.displayValue;

    // We don't render case numbers (per requirement), but we do a quick sanity check
    // so malformed payloads don't get counted under a real bucket.
    const looksLikeCase = typeof caseNumber === "string" && /^CS\d+/i.test(caseNumber);

    if (!looksLikeCase || typeof priorityDisplay !== "string") {
      counts.Unknown += 1;
      continue;
    }

    if (Object.prototype.hasOwnProperty.call(counts, priorityDisplay)) {
      counts[priorityDisplay] += 1;
    } else {
      counts.Unknown += 1;
    }
  }

  return counts;
}

function computePriorityCountsFromListLayout(graphqlResponseData) {
  const counts = {
    "1 - Critical": 0,
    "2 - High": 0,
    "3 - Moderate": 0,
    "4 - Low": 0,
    Unknown: 0
  };

  const rows =
    graphqlResponseData?.data?.GlideListLayout_Query?.getListLayout?.layoutQuery?.queryRows || [];

  for (const row of rows) {
    // Only handle non-grouped rows
    const rowData = row?.rowData;
    if (!Array.isArray(rowData)) {
      counts.Unknown += 1;
      continue;
    }

    const priorityCell = rowData.find((c) => c?.columnName === "priority");
    const priorityDisplay = priorityCell?.columnData?.displayValue;
    if (typeof priorityDisplay !== "string") {
      counts.Unknown += 1;
      continue;
    }

    if (Object.prototype.hasOwnProperty.call(counts, priorityDisplay)) {
      counts[priorityDisplay] += 1;
    } else {
      counts.Unknown += 1;
    }
  }

  return counts;
}

function formatCountsBlock(title, counts) {
  const lines = [`<strong>${title}</strong>`];
  for (const bucket of PRIORITY_BUCKETS) {
    lines.push(`${bucket} : ${counts[bucket]}`);
  }
  if (counts.Unknown) lines.push(`Unknown : ${counts.Unknown}`);
  return lines.join("\n");
}

function sendGraphql(body) {
  return new Promise((resolve) => {
    chrome.runtime.sendMessage({ type: "FETCH_QUEUE_JSON", body }, (response) => resolve(response));
  });
}

function buildLiveOpenCasesGraphqlBody() {
  // Mirroring the list page from the attached HAR (list-id 0e66..., tiny-id QF9k...)
  return {
    operationName: "nowRecordListConnected_min",
    variables: {
      table: "sn_customerservice_case",
      view: "",
      columns: "number,priority",
      fixedQuery: "",
      query:
        "active=true^assigned_toISEMPTY^stateNOT IN6,3,7^assignment_group.nameSTARTSWITHManu - Manu^ORassignment_group.nameSTARTSWITHManu - Pro",
      limit: 100,
      offset: 0,
      queryCategory: "list",
      maxColumns: 50,
      listId: "0e66f167977c4a184b77ff21f053af1f",
      listTitle: "Cases%20%20Manu%20-%20Manu%20All",
      runHighlightedValuesQuery: false,
      menuSelection: "sys_ux_my_list",
      ignoreTotalRecordCount: false,
      columnPreferenceKey: "",
      tiny: "QF9ktNm6N38E3iy4FZfKtZ5tr1DvDUPp"
    },
    query: `
      query nowRecordListConnected_min(
        $columns:String
        $listId:String
        $maxColumns:Int
        $limit:Int
        $offset:Int
        $query:String
        $fixedQuery:String
        $table:String!
        $view:String
        $runHighlightedValuesQuery:Boolean!
        $tiny:String
        $queryCategory:String
        $listTitle:String
        $menuSelection:String
        $ignoreTotalRecordCount:Boolean
        $columnPreferenceKey:String
      ){
        GlideListLayout_Query{
          getListLayout(
            columns:$columns
            listId:$listId
            maxColumns:$maxColumns
            limit:$limit
            offset:$offset
            query:$query
            fixedQuery:$fixedQuery
            table:$table
            view:$view
            runHighlightedValuesQuery:$runHighlightedValuesQuery
            tiny:$tiny
            queryCategory:$queryCategory
            listTitle:$listTitle
            menuSelection:$menuSelection
            ignoreTotalRecordCount:$ignoreTotalRecordCount
            columnPreferenceKey:$columnPreferenceKey
          ){
            layoutQuery{
              count
              queryRows{
                ... on GlideListLayout_QueryRowType{
                  rowData{columnName columnData{displayValue value}}
                }
              }
            }
          }
        }
      }
    `
  };
}

function buildGraphqlBody(condition) {
  return {
    operationName: "snPolarisLayout",
    variables: {
      condition
    },
    query: `
      query snPolarisLayout($condition:String){
        GlideRecord_Query{
          ui_notification_inbox(
            queryConditions:$condition
          ){
            _rowCount
            _results{
              sys_id{value}
              sys_created_on{value}
              payload{value}
              target_table{value}
              target{value}
              status{value}
              route{value}
              triggered_by{value displayValue}
            }
          }
        }
      }
    `
  };
}

async function run() {
  // Queue 1: My cases - Live (inbox notifications)
  // ServiceNow supports server-side javascript: in encoded queries (seen in the HAR
  // using gs.beginningOfLast7Days()). Using gs.getUserID() avoids needing to read
  // any page globals (which is blocked by strict CSP).
  const myCasesCondition = [
    "recipient=javascript:gs.getUserID()",
    "target_table=sn_customerservice_case",
    "sys_created_on>=javascript:gs.beginningOfLast7Days()",
    "status!=deleted",
    "ORDERBYDESCsys_created_on"
  ].join("^");

  const myCasesResponse = await sendGraphql(buildGraphqlBody(myCasesCondition));
  if (!myCasesResponse) {
    output.textContent = JSON.stringify({ error: "No response from background" }, null, 2);
    return;
  }

  if (myCasesResponse.needsLogin) {
    output.textContent = JSON.stringify(
      {
        error: "Please log in to https://support.ifs.com",
        hint: "Open support.ifs.com in a tab, sign in, then reopen this popup.",
        details: myCasesResponse.http || undefined
      },
      null,
      2
    );
    return;
  }

  if (!myCasesResponse.success) {
    output.textContent = JSON.stringify(myCasesResponse, null, 2);
    return;
  }

  const myCasesCounts = computePriorityCounts(myCasesResponse.data);

  // Queue 2: Live Open cases (list-id 0e66...)
  const liveOpenResponse = await sendGraphql(buildLiveOpenCasesGraphqlBody());
  if (!liveOpenResponse) {
    output.textContent = JSON.stringify({ error: "No response from background (live open cases)" }, null, 2);
    return;
  }

  if (liveOpenResponse.needsLogin) {
    output.textContent = JSON.stringify(
      {
        error: "Please log in to https://support.ifs.com",
        hint: "Open support.ifs.com in a tab, sign in, then reopen this popup.",
        details: liveOpenResponse.http || undefined
      },
      null,
      2
    );
    return;
  }

  if (!liveOpenResponse.success) {
    output.textContent = JSON.stringify(liveOpenResponse, null, 2);
    return;
  }

  const liveOpenCounts = computePriorityCountsFromListLayout(liveOpenResponse.data);

  output.innerHTML = [
    formatCountsBlock("My cases - Live", myCasesCounts),
    "",
    formatCountsBlock("Live Open cases", liveOpenCounts)
  ].join("\n");
}

run();
