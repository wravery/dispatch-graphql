// Get the default store and inbox folder IDs, and subscribe to updates on the most recent 10 items.
var subscriptionId = 0;
var subscriptionPayloads = [];
window.chrome.webview.hostObjects.graphql.fetchQuery(`
  query DefaultInboxIds {
    stores @orderBy(sorts: [
      {
      property: {id: 13312},
      type: BOOL,
      descending: true
      }
    ]) @take(count: 1) {
      name
      id
      specialFolders(ids: [INBOX]) {
        name
        id
        specialFolder
      }
    }
  }
`, "", "", (next) => {
    console.log(`Unexpected next callback from query:`);
    console.log(JSON.parse(next));
  })
  .then((payload) => {
    let { results } = JSON.parse(payload);
    console.log(`DefaultInboxIds:`);
    console.log(results);
    let store = results.data.stores[0];
    let inbox = store.specialFolders[0];
    return {
      storeId: store.id,
      objectId: inbox.id
    };
  })
  .then((variables) => {
    window.chrome.webview.hostObjects.graphql.fetchQuery(`
      subscription InboxItemsSubscription($storeId: ID!, $objectId: ID!) {
        items(folderId: {storeId: $storeId, objectId: $objectId}) @take(count: 10) {
          ... on ItemAdded {
            index
            added {
              ...ItemFragment
            }
          }
          ... on ItemUpdated {
            index
            updated {
              ...ItemFragment
            }
          }
          ... on ItemRemoved {
            index
            removed
          }
          ... on ItemsReloaded {
            reloaded {
              ...ItemFragment
            }
          }
        }
      }
    
      fragment ItemFragment on Item {
        id
        subject
        read
        received
        modified
        sender
        to
        cc
        preview
      }
    `, "", JSON.stringify(variables), (payload) => {
      let next = JSON.parse(payload);
      subscriptionPayloads.push(next);
      console.log(next);
    })
    .then((results) => {
      let {pending} = JSON.parse(results);
      console.log(`Pending subscription: ${pending}`);
      subscriptionId = pending;
    });
  });
