

```mermaid
%%{
  init: {
    'theme': 'base',
    'themeVariables': {
      'primaryColor': '#F4F4F4',
      'primaryTextColor': '#212121',
      'lineColor': '#555',
      'secondaryColor': '#B0C4DE',
      'tertiaryColor': '#FFFFFF'
    }
  }
}%%
graph TD
    subgraph A[Standard Workflow (PAM Available)]
        A1(Incident Occurs<br>Emergency Access Needed<br>[IT Ops])
        A1 --> A2{Submit Request<br>via PAM System<br>[Requestor]}
        A2 --> A3(Approve Access Request<br>[InfoSec])
        A3 -- Approved --> A4(Get JIT Access<br>via PAM<br>[Requestor])
        A3 -- Denied --> A5(Process Ends)
        A4 --> Z1
    end

    subgraph B[Disaster Workflow (PAM Unavailable)]
        B1(Major Disaster<br>PAM is Down<br>[Leadership/InfoSec])
        B1 --> B2{Fill Out<br>Emergency Form<br>[Requestor]}
        B2 --> B3(Dual-Executive Approval<br>(e.g., CISO + IT Director)<br>[Leadership])
        B3 -- Approved --> B4(Retrieve Sealed Password<br>from Physical Safe<br>[Leadership])
        B4 --> B5(Perform Manual Login<br>Under Supervision<br>[Requestor/InfoSec])
        B3 -- Denied --> B6(Process Ends)
        B5 --> Z1
    end

    Z1(Execute Necessary Tasks<br>[Requestor/IT Ops])
    Z1 --> Z2((SIEM Alert Triggered<br>on Successful Login<br>[SOC/SIEM]))
    Z2 --> Z3{Task Complete<br>Reset Password Immediately<br>[IT Ops]}
    Z3 --> Z4(Convene eCAB<br>for Post-Incident Review<br>[eCAB])
    Z4 --> Z5(Generate Audit Report<br>and Archive<br>[InfoSec])

    style A1 fill:#FFDDC1,stroke:#333,stroke-width:2px
    style B1 fill:#FFB3B3,stroke:#333,stroke-width:2px,font-weight:bold
    style Z2 fill:#FFFFB3,stroke:#333,stroke-width:2px
    style Z4 fill:#C1FFC1,stroke:#333,stroke-width:2px
