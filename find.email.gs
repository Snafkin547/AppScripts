/**
 * Search emails with subject containing a keyword
 * * @param {string} keyword - any keyword you wish emails to contain
 * * @param {int} numThreads (optional) - max number of threads
 * * @param {int} max_length (optional) - max length of truncated email chain
*/

function find_emails(keyword, numThreads = 20 ,max_length = 50000){
  const query = `"${keyword}"`;
  const threads = GmailApp.search(query, 0, numThreads);

  if (threads.length==0){
    throw new Error(`No case-related email found for ${keyword}.`);
  } 
  else {
    let content = '';
    threads.forEach(thread => {
      const subject = thread.getFirstMessageSubject();
      if (subject.includes(keyword)) {
        content += `\n\n--- THREAD START: ${subject} ---\n`;
        thread.getMessages().forEach(message => {
          content += `--- Email Start (From: ${message.getFrom()}, Date: ${message.getDate()}) ---\n${message.getPlainBody()}\n--- Email End ---\n\n`;
        });
      }
    });

    if (content === '') {
      throw new Error(`Emails found, but none with the keyword: ${keyword} in the subject.`);
    }

    const truncationMessage = "\n... (Email content truncated)";
    return content.length > max_length ? content.substring(0, max_length) + truncationMessage : content;
  }
}