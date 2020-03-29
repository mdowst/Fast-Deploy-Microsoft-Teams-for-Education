
# ----------------------------------------------------------------------------------
#
# Copyright Microsoft Corporation
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
# http://www.apache.org/licenses/LICENSE-2.0
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
# ----------------------------------------------------------------------------------

<#
.Synopsis
Submit a new batch of policy assignments
.Description
Submit a new batch of policy assignments
.Example
PS C:\> {{ Add code here }}

{{ Add output here }}
.Example
PS C:\> {{ Add code here }}

{{ Add output here }}

.Link
https://docs.microsoft.com/en-us/powershell/module/microsoft.teams.config/new-csbatchpolicyassignmentoperation
#>
function New-CsBatchPolicyAssignmentOperation {
[OutputType([System.String])]
[CmdletBinding(DefaultParameterSetName='NewExpanded', PositionalBinding=$false, SupportsShouldProcess, ConfirmImpact='Medium')]
param(
    [Parameter()]
    [Microsoft.Teams.Config.Cmdlets.Category('Query')]
    [System.String]
    # string
    ${OperationName},

    [Parameter(ParameterSetName='New', Mandatory, ValueFromPipeline)]
    [Microsoft.Teams.Config.Cmdlets.Category('Body')]
    [Microsoft.Teams.Config.Cmdlets.Models.IBatchAssignBody]
    # .
    # To construct, see NOTES section for PAYLOAD properties and create a hash table.
    ${Payload},

    [Parameter(ParameterSetName='NewExpanded', Mandatory)]
    [Alias('User')]
    [Microsoft.Teams.Config.Cmdlets.Category('Body')]
    [System.String[]]
    # .
    ${Identity},

    [Parameter(ParameterSetName='NewExpanded', Mandatory)]
    [Microsoft.Teams.Config.Cmdlets.Category('Body')]
    [System.String]
    # .
    ${PolicyName},

    [Parameter(ParameterSetName='NewExpanded', Mandatory)]
    [Microsoft.Teams.Config.Cmdlets.Category('Body')]
    [System.String]
    # .
    ${PolicyType},

    [Parameter(ParameterSetName='NewExpanded')]
    [Microsoft.Teams.Config.Cmdlets.Category('Body')]
    [Microsoft.Teams.Config.Cmdlets.Runtime.Info(PossibleTypes=([Microsoft.Teams.Config.Cmdlets.Models.IBatchAssignBodyAdditionalParameters]))]
    [System.Collections.Hashtable]
    # .
    ${AdditionalParameters},

    [Parameter(DontShow)]
    [Microsoft.Teams.Config.Cmdlets.Category('Runtime')]
    [System.Management.Automation.SwitchParameter]
    # Wait for .NET debugger to attach
    ${Break},

    [Parameter(DontShow)]
    [ValidateNotNull()]
    [Microsoft.Teams.Config.Cmdlets.Category('Runtime')]
    [Microsoft.Teams.Config.Cmdlets.Runtime.SendAsyncStep[]]
    # SendAsync Pipeline Steps to be appended to the front of the pipeline
    ${HttpPipelineAppend},

    [Parameter(DontShow)]
    [ValidateNotNull()]
    [Microsoft.Teams.Config.Cmdlets.Category('Runtime')]
    [Microsoft.Teams.Config.Cmdlets.Runtime.SendAsyncStep[]]
    # SendAsync Pipeline Steps to be prepended to the front of the pipeline
    ${HttpPipelinePrepend},

    [Parameter(DontShow)]
    [Microsoft.Teams.Config.Cmdlets.Category('Runtime')]
    [System.Uri]
    # The URI for the proxy server to use
    ${Proxy},

    [Parameter(DontShow)]
    [ValidateNotNull()]
    [Microsoft.Teams.Config.Cmdlets.Category('Runtime')]
    [System.Management.Automation.PSCredential]
    # Credentials for a proxy server to use for the remote call
    ${ProxyCredential},

    [Parameter(DontShow)]
    [Microsoft.Teams.Config.Cmdlets.Category('Runtime')]
    [System.Management.Automation.SwitchParameter]
    # Use the default credentials for the proxy
    ${ProxyUseDefaultCredentials}
)

begin {
    try {
        $outBuffer = $null
        if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer)) {
            $PSBoundParameters['OutBuffer'] = 1
        }
        $parameterSet = $PSCmdlet.ParameterSetName
        $mapping = @{
            New = 'Microsoft.Teams.Config.private\New-CsBatchPolicyAssignmentOperation_New';
            NewExpanded = 'Microsoft.Teams.Config.private\New-CsBatchPolicyAssignmentOperation_NewExpanded';
        }
        $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand(($mapping[$parameterSet]), [System.Management.Automation.CommandTypes]::Cmdlet)
        $scriptCmd = {& $wrappedCmd @PSBoundParameters}
        $steppablePipeline = $scriptCmd.GetSteppablePipeline($MyInvocation.CommandOrigin)
        $steppablePipeline.Begin($PSCmdlet)
    } catch {
        throw
    }
}

process {
    try {
        $steppablePipeline.Process($_)
    } catch {
        throw
    }
}

end {
    try {
        $steppablePipeline.End()
    } catch {
        throw
    }
}
}
