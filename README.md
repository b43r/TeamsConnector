# About TeamsConnector
A simple package for interfacing with Microsoft Teams, written in C# and .NET 6.0.

Main features:
- Get presence of current user
- Get notified about presence changed
- Get notified about incoming calls, including the calling phone number
- Make outgoing calls

Works with both the old and **new (beta)** Teams client!


# Getting started

Create a new instance of TeamsClient and get presence of the current user:

```csharp
var teamsClient = new TeamsClient();
string presence = teamsClient.GetAvailability() + "/" + teamsClient.GetActivity();
```

Get notified about presence changes:

```csharp
teamsClient.PresenceChanged += TeamsClient_OnPresenceChanged;
teamsClient.CreatePresenceSubscription();
```

```csharp
private void TeamsClient_OnPresenceChanged(object? sender, PresenceChangedEventArgs e)
{
  Console.WriteLine($"Presence changed: {e.Availability}/{e.Activity}");
}
```
