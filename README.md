# About TeamsConnector
A simple package for interfacing with Microsoft Teams, written in C# and .NET 6.0.

Main features:
- Get presence of current user
- Get notified about presence changed
- Get notified about incoming calls, including the calling phone number
- Make outgoing calls

Works with both the old and **new (beta)** Teams client!


## Getting started

Create a new instance of TeamsClient and get presence of the current user:

```csharp
var teamsClient = new TeamsClient();
string presence = teamsClient.GetAvailability() + "/" + teamsClient.GetActivity();
```

Get notified about presence changes:

```csharp
teamsClient.PresenceChanged += TeamsClient_PresenceChanged;
teamsClient.CreatePresenceSubscription();
```

```csharp
private void TeamsClient_PresenceChanged(object? sender, PresenceChangedEventArgs e)
{
    Console.WriteLine($"Presence changed: {e.Availability}/{e.Activity}");
}
```

Get notified about incoming calls:

```csharp
teamsClient.IncomingCall += TeamsClient_IncomingCall;
```

```csharp
private void TeamsClient_IncomingCall(object? sender, IncomingCallEventArgs e)
{
    Console.WriteLine($"Incoming call from: {e.PhoneNumber}");
}
```

Start an outgoing call:

```csharp
teamsClient.MakeCall("+41441112233");
```

## Error handling

The constructor of the TeamsClient class throws a ```TeamsConnectorException``` exception if:
- Microsoft Teams ist not registered as the default IM application
- MIcrosoft Teams is not started
