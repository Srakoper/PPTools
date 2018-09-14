/**
 * Stores OPs and names of accounts to have labels removed.
 * Accounts to have labels removed must be manually added.
 * If account to have label removed does not contain the label in question, it will be ignored when processing.
 */
var accountsToEnd = {OP0712341: "Izpuh center, d.o.o.",
                     OP0760485: "Digital Logic, d.o.o."};

/**
 * Fetches OP from campaign name(s) by iterating over campaigns in campaign iterator.
 * OP format: OP + 7 digits (length: 9).
 * Requires OP to be present at the begining of campaign name(s).
 * If no OP is found, 'N/A' is returned.
 * @param campaignIterator: campaign iterator to check campaign names in
 * @return: str; OP or 'N/A' if no OP found
 */
function getOP(campaignIterator) {
  while (campaignIterator.hasNext()) {
      var OP = campaignIterator.next().getName().substring(0, 9);
      if (OP.substring(0,2) == "OP") return OP;
      else if (!(campaignIterator.hasNext())) return "N/A";
  }
}

/*
 * Removes specified label from accounts that contain it.
 * To mark account as "ended" (not being processed by other scripts), pass label "Aktivne" as argument to remove.
 * @param label: str; valid label for selecting accounts containing it
 * @param toRemove: str; valid label to remove from selected accounts
 * @param mapOfAccounts: object; map of OP : account name pairs of accounts to be processed
 */
function removeLabel(label, toRemove, mapOfAccounts) {
  var accountIterator = MccApp.accounts().withCondition('LabelNames CONTAINS "' + label + '"').get();
  while (accountIterator.hasNext()) {
    var account = accountIterator.next();
    Logger.log(account.getName());
    MccApp.select(account);
    var OP = getOP(AdWordsApp.campaigns().get());
    if (OP in mapOfAccounts) {
      var accountLabelIterator = MccApp.accountLabels().get();
      while (accountLabelIterator.hasNext()) {
        var accountLabel = accountLabelIterator.next();
        if (accountLabel.getName() === toRemove) {
          account.removeLabel(accountLabel.getName());
          Logger.log("Label " + toRemove + " removed from account " + OP + " " + mapOfAccounts[OP]);
        }
      }
    }
  }
}

function main() {
  removeLabel("Aktivne", "Aktivne", accountsToEnd);
}
