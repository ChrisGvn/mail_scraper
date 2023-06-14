  <h2>Briefly</h2>
  <p>A Python script that scrapes an Outlook folder full of automated logs. These logs are to be summarized into an easier-to-read form.</p>

  <h2>Specifically</h2>
  <p>A custom -let's call it- monitoring system produces logs about the status of an application running on multiple servers and sends these logs as emails.</p>
  <p>There is no information in the body of the emails, all we're interested in is the title of the email, and it resembles the following:</p>
  <pre><code>RemoteSite.SERVERNAME: Application status changed (online).
  or
RemoteSite.SERVERNAME: Application status changed (offline).</code></pre>

  <p>I can't describe what the system is exactly, but disconnections and reconnections are normal and they happen quite a lot on a considerable number of machines. What we have to do is ensure that if a disconnection happens, a reconnection must follow within a short amount of time.</p>
  <p>The volume of all these emails makes the task gruesome and time-consuming, so the procedure is in dire need of automation.</p>

  <p><em> - This is a work in progress and has not been tested adequately -</em></p>
