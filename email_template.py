def get_email_subject():
    return "SRE / DevOps Engineer | 3+ Years Experience"


def get_email_body(company):
    if not company or str(company).strip() == "":
        greeting = "Hi Hiring Team,"
    else:
        greeting = f"Hi {str(company).strip()} Team,"

    body = f"""
{greeting}

I'm Nishant Mandil, a Site Reliability Engineer at Colt Technology Services, currently focused on improving reliability and automation in large production environments.

In the past few years I've worked on problems like:

- Improving system uptime by 20% through reliability engineering
- Reducing incident response time by 30% with better monitoring and alerting
- Building CI/CD pipelines using Jenkins, Docker, and Kubernetes
- Automating operational workflows using Python and AWS

Alongside this, I've been actively exploring AI applications in SRE workflows. Recently I built an AI-assisted log analyzer that helps engineers identify root causes faster during production incidents.

The project uses LLM-based analysis to detect patterns in logs and generate incident insights automatically, helping reduce the time engineers spend manually investigating failures.

GitHub Project:
https://github.com/nishantmandil/sre-log-analyzer

I'm now looking to apply these skills in a DevOps/SRE role where reliability, automation, and intelligent tooling truly matter.

If your team is working on interesting infrastructure or platform problems, I'd love to connect.

Best regards,
Nishant Mandil
📍 Gurugram, India
📞 +91 8085569375
📧 nishantmandil105@gmail.com

LinkedIn:
https://www.linkedin.com/in/nishant-mandil-07b165159/

GitHub:
https://github.com/nishantmandil

Portfolio:
https://portfolio.taurbykaur.co.in/
"""
    return body