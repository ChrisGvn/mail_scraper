<!DOCTYPE html>
<html>
<head>
  <title>README</title>
  <style>
    body {
      font-family: Arial, sans-serif;
      margin: 20px;
      line-height: 1.5;
    }

    h2 {
      font-size: 20px;
      margin-top: 20px;
    }

    code {
      font-family: Consolas, monospace;
      background-color: #f8f8f8;
      padding: 2px 4px;
      border: 1px solid #ddd;
      border-radius: 4px;
    }

    pre {
      font-family: Consolas, monospace;
      background-color: #f8f8f8;
      padding: 10px;
      border: 1px solid #ddd;
      border-radius: 4px;
      overflow-x: auto;
    }
  </style>
</head>
<body>
  <h2>Briefly</h2>
  <p>A Python script that scrapes an Outlook folder full of automated logs. These logs are to be summarized into an easier-to-read form.</p>

  <h2>Specifically</h2>
  <p>A custom -let's call it- monitoring system produces logs about the status of an application running on multiple servers and sends these logs as emails.</p>
  <p>There is no information in the body of the emails, all we're interested in is the title of the email, and it resembles the following:</p>
  <pre><code>RemoteSite.SERVERNAME: Application status changed (online).
  or
RemoteSite.SERVERNAME: Application status changed (offline).</code></pre>

  <p>I can't describe what the system is exactly, but disconnections and reconnections are normal and they happen quite a lot on a considerable number of machines.</p>
  <p>What we have to do is ensure that if a disconnection happens, a reconnection must follow within a short amount of time.</p>
  <p>The volume of all these emails makes the task gruesome and time-consuming, so the procedure is in dire need of automation.</p>

  <p><em> - This is a work in progress and has not been tested adequately -</em></p>
</body>
</html>
