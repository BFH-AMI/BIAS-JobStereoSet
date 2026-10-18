# The BIAS JobStereoSet Benchmark

To support the evaluation of stereotypical associations in multilingual language models, the BIAS project developed JobStereoSet, a curated stereotype evaluation dataset covering five consortium languages: German (DE), Dutch (NL), Italian (IT), Turkish (TK), and Norwegian (NO).

## Acknowledgements
This work is part of the Europe Horizon project BIAS funded by the European Commission, and has received funding from the Swiss State Secretariat for Education, Research and Innovation (SERI). All work from the BIAS Project: https://github.com/BFH-AMI/BIAS 

The work in this repository is inspired by the data structure and approach of the following paper:
Moin Nadeem, Anna Bethke, and Siva Reddy. 2020. Stereoset: Measuring stereotypical bias in pretrained language models. https://aclanthology.org/2021.acl-long.416/

## Dataset Description
The dataset builds on the general approach of StereoSet (Nadeem et al, 2020), while focusing specifically on stereotypes related to employment and the labor market and covering European languages. Rather than directly translating existing English-language benchmarks, JobStereoSet incorporates language-specific expressions and concepts identified through the BIAS project's co-creation activities.

The dataset was developed through a multilingual survey conducted in collaboration with native speakers from the consortium and using a crowdworker platform. The collected responses were reviewed by native speakers to assess their linguistic and contextual appropriateness. Subsequently, the data was processed through an interdisciplinary workflow combining manual review and automated processing steps. The code of the automated processing steps is documented in this repository. The target expressions used in the dataset, such as for example working mother and working father in the German dataset, were informed by the outcomes of the BIAS co-creation activities, linking the evaluation material to the project's broader investigation of stereotypes in employment contexts.

The resulting dataset consists of sentence templates containing a blank position, accompanied by three possible completions: a stereotypical, an anti-stereotypical, and an unrelated (random) option. These examples were crafted by the participants of the survey. This structure allows language models to be evaluated by prompting them to select among the available completions and examining their preferences for stereotypical versus anti-stereotypical associations. JobStereoSet thus provides a structured basis for comparing stereotype-related model behavior across the included languages and for assessing the effects of bias mitigation approaches in multilingual NLP applications.
