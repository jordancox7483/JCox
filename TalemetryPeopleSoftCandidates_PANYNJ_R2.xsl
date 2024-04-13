<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <xsl:output cdata-section-elements="TextResume"/>
  <xsl:key name="questionGroup" match="//ResumeContents/ResumeFormats/OtherExtractedResults/CustomFields/CustomField" use="substring-before(name, '_')"/>
  <!-- 
  The input to this transform is expected to be the Transaction node of an
  extracted resume from Talemetry Apply, with default namespace explicitly
  removed to simplify xpath query (default ns: http://ns.hr-xml.org)
  -->
  <xsl:template match="/">
    <SOAP-ENV:Envelope xmlns="" xmlns:SOAP-ENV="http://schemas.xmlsoap.org/soap/envelope/" xmlns:SOAP-ENC="http://schemas.xmlsoap.org/soap/encoding/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
      <SOAP-ENV:Body>
        <m:Input xmlns:m="http://peoplesoft.com/HCM/Schema">
          <Candidate>
          <Candidate xmlns="http://ns.hr-xml.org">
            <RelatedPositionPostings>
              <xsl:for-each select="//Transaction/JobCode">
                <xsl:call-template name="jobcodeTokenize">
                  <xsl:with-param name="str" select="."/>
                </xsl:call-template>
              </xsl:for-each>
            </RelatedPositionPostings>


				
			
            <Resume xmlns:hr="http://ns.hr-xml.org/PersonDescriptors" xmlns:rm="resumemirror.com">
              <xsl:attribute name="xml:lang">
                <xsl:value-of select="//Resume/@xml:lang" /> 
              </xsl:attribute>
              <StructuredXMLResume>
                <xsl:for-each select="//Resume/StructuredXMLResume/ContactInfo">
                  <ContactInfo>
                    <xsl:apply-templates select="@*|node()"/>
                  </ContactInfo>
                </xsl:for-each>
		
				<xsl:template match="ContactMethod[contains(Use,'personal')]/PostalAddress">
				  <PostalAddress>
					<xsl:apply-templates select="@*|node()"/>
					<Region><xsl:value-of select="//CustomFields/CustomField[contains(name,'Question-ADR_County')]/value"/></Region>
				  </PostalAddress>
				</xsl:template>
				
				
                <xsl:for-each select="//Resume/StructuredXMLResume/EmploymentHistory">
                  <EmploymentHistory>
                    <xsl:apply-templates select="@*|node()"/>
                  </EmploymentHistory>
                </xsl:for-each>
                <xsl:for-each select="//Resume/StructuredXMLResume/EducationHistory">
                  <EducationHistory>
                    <xsl:apply-templates select="@*|node()"/>
                  </EducationHistory>
                </xsl:for-each>
                <LicensesAndCertifications>
                  <xsl:for-each select="//Resume/StructuredXMLResume/Qualifications/Competency[contains(TaxonomyId/@id,'PS89_LicenseCert') and not(@name=preceding-sibling::Competency/@name)]">
                    <LicenseOrCertification>
                      <Name>
                        <xsl:value-of select="@name"/>
                      </Name>                          
                    </LicenseOrCertification>
                  </xsl:for-each>
                </LicensesAndCertifications>
                <Associations>
                  <xsl:for-each select="//Resume/StructuredXMLResume/Qualifications/Competency[contains(TaxonomyId/@id,'PS89_Membership') and not(@name=preceding-sibling::Competency/@name)]">
                    <Association>
                      <Name>
                        <xsl:value-of select="@name"/>
                      </Name>
                    </Association>                     
                  </xsl:for-each>
                </Associations>
                <Languages>
                  <xsl:for-each select="//Resume/StructuredXMLResume/Qualifications/Competency[contains(TaxonomyId/@id,'PS89_Language') and not(@name=preceding-sibling::Competency/@name)]">
                    <Language>
                      <LanguageCode>
                        <xsl:value-of select="@name"/>
                      </LanguageCode>
                      <Read>0</Read>
                      <Write>0</Write>
                      <Speak>0</Speak>
                    </Language>
                  </xsl:for-each>
                </Languages>
                <Qualifications>
                  <xsl:for-each select="//Resume/StructuredXMLResume/Qualifications/Competency[contains(TaxonomyId/@id,'PS89_Competency') and not(@name=preceding-sibling::Competency/@name)]">
                    <Competency>
                      <xsl:apply-templates select="@*|node()"/>
                    </Competency>
                  </xsl:for-each>
                </Qualifications>
                <xsl:for-each select="//Resume/StructuredXMLResume/References">
                  <References>
                    <xsl:apply-templates select="@*|node()"/>
                  </References>
                </xsl:for-each>
                <xsl:for-each select="//Resume/StructuredXMLResume/RevisionDate">
                  <RevisionDate>
                    <xsl:apply-templates select="@*|node()"/>
                  </RevisionDate>
                </xsl:for-each>
              </StructuredXMLResume>
              <NonXMLResume>
                <xsl:for-each select="//ResumeContents/ResumeFormats/TextResume">
                  <TextResume>
                    <xsl:apply-templates select="@*|node()"/>
                  </TextResume>
                </xsl:for-each>
              </NonXMLResume>
            </Resume>
            <xsl:for-each select="//Resume/UserArea">
              <UserArea xmlns:hr="http://ns.hr-xml.org/PersonDescriptors" xmlns:rm="resumemirror.com">
                <xsl:apply-templates select="@*|node()"/>
                <xsl:apply-templates select="//ResumeContents/ResumeFormats/OtherExtractedResults/CustomFields"/>
                <SubAccountUserID>
                  <xsl:value-of select="//Transaction/SubAccountUserID"/>
                </SubAccountUserID>
                <!--
                This was done in one of the customized XSLs, but seems to be unneeded.
                <xsl:call-template name="removeNamespacePrefix">
                  <xsl:with-param name="nodes" select="rm:Assessments"/>
                </xsl:call-template>
                -->
              </UserArea>
            </xsl:for-each>
            <RMCustomFields>
              <BuildSource>EAM</BuildSource>
              <GlobalAutoReply>N</GlobalAutoReply>
            </RMCustomFields>
          </Candidate>
          </Candidate>
          <Attachment>
            <Type>Resume</Type>
            <Name>
              <xsl:value-of select="//ResumeContents/ResumeFormats/Base64BinaryResume/@filename"/>
            </Name>
            <File>
              <xsl:value-of select="//ResumeContents/ResumeFormats/Base64BinaryResume"/>
            </File>
          </Attachment>
          <xsl:for-each select="//Resume/UserArea/rm:AdditionalAttachments/rm:Attachment[@valid='true']" xmlns:rm="resumemirror.com">
            <Attachment>
              <Type>Other</Type>
              <Name>
                <xsl:value-of select="rm:filename"/>
              </Name>
              <File>
                <xsl:value-of select="concat('{{attachment_id{{', rm:id, '}}}}')"/>
              </File>
            </Attachment>
          </xsl:for-each>
        </m:Input>
      </SOAP-ENV:Body>
    </SOAP-ENV:Envelope>
  </xsl:template>
  <!-- 
  An identity transform for copying nodes recursively. This
  would be overridden for nodes that need special handling.
  -->
  <xsl:template match="*">
    <xsl:element name="{local-name()}">
      <xsl:apply-templates select="@*|node()"/>
    </xsl:element>
  </xsl:template>
  <xsl:template match="m:*|hr:*|rm:*" xmlns:m="http://peoplesoft.com/HCM/Schema" xmlns:hr="http://ns.hr-xml.org/PersonDescriptors" xmlns:rm="resumemirror.com">
    <xsl:copy>
      <xsl:apply-templates select="@*|node()"/>
    </xsl:copy>
  </xsl:template>
  <xsl:template match="@*|comment()|text()|processing-instruction()">
    <xsl:copy>
      <xsl:apply-templates select="@*|node()"/>
    </xsl:copy>
  </xsl:template>
  <!--	JC - 2017 - Set the variable to the current year here for use below when the end date is "current"  -->
  <xsl:variable name="current_year" select="substring(Transaction/TransactionDateTime, 5,4)"/>
  
  
  
  <!--
  Override identity transform to handle current date in education.
  -->
  <xsl:template match="//EducationHistory/SchoolOrInstitution/Degree/DegreeDate/StringDate |
                       //EducationHistory/SchoolOrInstitution/Degree/DatesOfAttendance/StartDate/StringDate |
                       //EducationHistory/SchoolOrInstitution/Degree/DatesOfAttendance/EndDate/StringDate">
    <xsl:choose>
      <xsl:when test=".='current'">
        <Year>
          <xsl:value-of select="$current_year"/>
        </Year>
      </xsl:when>
      <xsl:otherwise>
        <StringDate>
          <xsl:apply-templates select="@*|node()"/>
        </StringDate>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!--
  Override identity transform to handle current date in education.
  -->
  <xsl:template match="//EducationHistory/SchoolOrInstitution/UserArea/rm:DatesOfAttendance/rm:StartDate/rm:StringDate |
                       //EducationHistory/SchoolOrInstitution/UserArea/rm:DatesOfAttendance/rm:EndDate/rm:StringDate" xmlns:rm="resumemirror.com">
    <xsl:choose>
      <xsl:when test=".='current'">
        <rm:Year>
          <xsl:value-of select="$current_year"/>
        </rm:Year>
      </xsl:when>
      <xsl:otherwise>
        <StringDate>
          <xsl:apply-templates select="@*|node()"/>
        </StringDate>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!--
  Override identity transform to names of direct apply question fields.
  -->
  



  <xsl:template match="//ResumeContents/ResumeFormats/OtherExtractedResults/CustomFields">
    <xsl:element name="Questions" namespace="">
      <xsl:for-each select="CustomField[generate-id()=generate-id(key('questionGroup', substring-before(name, '_'))[1])]">
        <xsl:sort select="name" order="ascending"/>
        <xsl:if test="string-length(substring-before(name, '_')) = 12">
          <xsl:variable name="group_name" select="substring(name, 10, 3)"/>
          <xsl:variable name="group_elem">
            <xsl:call-template name="questionGroupName">
              <xsl:with-param name="str" select="$group_name"/>
            </xsl:call-template>
          </xsl:variable>
          <xsl:element name="{$group_elem}">
            <xsl:for-each select="../CustomField[substring(name, 10, 3) = $group_name]">
              <xsl:sort select="name" order="ascending"/>
              <xsl:variable name="field_elem">
                <xsl:call-template name="questionFieldName">
                  <xsl:with-param name="str" select="substring-after(name, '_')"/>
                </xsl:call-template>
              </xsl:variable>
	
<!-- JC - 2017 PANYNJ - Added variable to replace "I decline" answer when it is chosen by applicant, XSLT 1.0 has no replace function  -->		
						  <xsl:variable name="myVar">
							<xsl:call-template name="string-replace-all">
							  <xsl:with-param name="text" select="value" />
							  <xsl:with-param name="replace" select="',&quot;SELFID&quot;'" />
							  <xsl:with-param name="by" select="''" />
							</xsl:call-template>
						  </xsl:variable>
			  
              <xsl:element name="{$field_elem}">
			  
<!-- JC - 2017 PANYNJ - Added Choose to block "I decline" answer when it is chosen by applicant  -->			  
				<xsl:choose>
				  <xsl:when test="name='Question-EEO_EthnicityUS' and contains(value,'SELFID')"> <xsl:value-of select="$myVar" /> </xsl:when>
				  <xsl:when test="name='Question-CST_EthnicityUS' and contains(value,'SELFID')"/>
				  <xsl:otherwise>				  
					<xsl:value-of select="value"/>
				  </xsl:otherwise>
				</xsl:choose>
			
              </xsl:element>
            </xsl:for-each>
          </xsl:element>
        </xsl:if>
      </xsl:for-each>
    </xsl:element>
    <xsl:element name="CustomFields" namespace="">
      <xsl:for-each select="CustomField">
        <xsl:sort select="name" order="ascending"/>
        <xsl:choose>
          <xsl:when test="string-length(substring-before(name, '_')) = 12"/>
          <xsl:when test="name='Source' and (value='TalemetryConnect' or value='Talent Community' or value='importer')"/>
          <xsl:when test="name='SubSource' and preceding-sibling::CustomField/name='Source' and (preceding-sibling::CustomField/value='TalemetryConnect' or preceding-sibling::CustomField/value='Talent Community' or preceding-sibling::CustomField/value='importer')"/>
          <xsl:when test="name='SubSource' and following-sibling::CustomField/name='Source' and (following-sibling::CustomField/value='TalemetryConnect' or following-sibling::CustomField/value='Talent Community' or following-sibling::CustomField/value='importer')"/>
		  
          <xsl:otherwise>
            <xsl:variable name="field_elem">
              <xsl:call-template name="customFieldName">
                <xsl:with-param name="str" select="name"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:element name="{$field_elem}">
              <xsl:value-of select="value"/>
            </xsl:element>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each>
      <ApplicationDateTime>
        <xsl:value-of select="//Transaction/TransactionDateTime"/>
      </ApplicationDateTime>
    </xsl:element>
  </xsl:template>
  <!--
  Function template for removing namespace of node and descendants. If this
  is applicable to all occurrences of the node, we may be able to just use
  a template match, such as:
  <xsl:template match="//Resume/UserArea/rm:Assessments |
                       //Resume/UserArea/rm:Assessments//*" xmlns:rm="resumemirror.com">
    <xsl:element name="{local-name()}">
      <xsl:apply-templates select="@*|node()"/>
    </xsl:element>
  </xsl:template>
  -->
  <xsl:template name="removeNamespacePrefix">
    <xsl:param name="nodes"/>
    <xsl:for-each select="$nodes">
      <xsl:element name="{local-name()}">
        <xsl:call-template name="removeNamespacePrefix">
          <xsl:with-param name="nodes" select="*"/>
        </xsl:call-template>
        <xsl:apply-templates select="@*|text()[normalize-space(.)]"/>
      </xsl:element>
    </xsl:for-each>
  </xsl:template>
  <!--
  Function template for tokenizing job codes into separate nodes.
  -->
  <xsl:template name="jobcodeTokenize">
    <xsl:param name="str"/>
    <xsl:if test="string-length($str)&gt;0">
      <xsl:variable name="prior-str">
        <xsl:choose>
          <xsl:when test="contains($str, ',')">
            <xsl:value-of select="substring-before($str, ',')"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$str"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:variable name="after-str">
        <xsl:choose>
          <xsl:when test="contains($str, ',')">
            <xsl:value-of select="substring-after($str, ',')"/>
          </xsl:when>
        </xsl:choose>
      </xsl:variable>
      <PositionPosting>
        <Id>
          <IdValue name="JobOpeningID">
            <xsl:call-template name="jobcodeStrip">
              <xsl:with-param name="str" select="$prior-str"/>
            </xsl:call-template>
          </IdValue>
        </Id>
      </PositionPosting>
      <xsl:call-template name="jobcodeTokenize">
        <xsl:with-param name="str" select="$after-str"/>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>
  

<!-- JC - 2017 - Function template for moving County. Looks for custom field configured in Talemetry as "ADR_County" and moves the data to Region2 for Peoplesoft mapping -->
<!-- Needed because using the custom mapping table generates orphaned rows -->		  
<xsl:template match="ContactMethod[contains(Use,'personal')]/PostalAddress">
  <PostalAddress>
	<xsl:apply-templates select="@*|node()"/>
	<Region><xsl:value-of select="//CustomFields/CustomField[contains(name,'Question-ADR_County')]/value"/></Region>
  </PostalAddress>
</xsl:template>



<!-- JC - 2017 - Function template to truncate the degree/name if longer than x characters  -->
<!-- Added to address cases where very long names caused resume to fail into peoplesoft  -->	  
    <xsl:template match="DegreeMajor/Name[string-length(normalize-space())>=60]/text()">
        <xsl:value-of select="substring(normalize-space(),1,59)"/>
    </xsl:template>


<!-- JC - 2017 PANYNJ - Added FN to replace "I decline" answer when it is chosen by applicant, XSLT 1.0 has no replace function  -->	
 <xsl:template name="string-replace-all">
    <xsl:param name="text" />
    <xsl:param name="replace" />
    <xsl:param name="by" />
    <xsl:choose>
      <xsl:when test="contains($text, $replace)">
        <xsl:value-of select="substring-before($text,$replace)" />
        <xsl:value-of select="$by" />
        <xsl:call-template name="string-replace-all">
          <xsl:with-param name="text"
          select="substring-after($text,$replace)" />
          <xsl:with-param name="replace" select="$replace" />
          <xsl:with-param name="by" select="$by" />
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$text" />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>



  <!--
  Function template for removing sequence id from job codes.
  -->
  <xsl:template name="jobcodeStrip">
    <xsl:param name="str"/>
    <xsl:choose>
      <xsl:when test="contains($str, '-')">
        <xsl:value-of select="normalize-space(substring-before($str, '-'))"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="normalize-space($str)"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!--
  Function template for cleansing names of direct apply question fields.
  -->
  <xsl:template name="customFieldName">
    <xsl:param name="str"/>
    <xsl:variable name="str_norm" select="normalize-space($str)"/>
    <xsl:choose>
      <xsl:when test="starts-with($str_norm, 'Question-')">
        <xsl:call-template name="elementName">
          <xsl:with-param name="str" select="substring-after($str_norm, 'Question-')"/>
          <xsl:with-param name="prefix" select="'Q'"/>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="elementName">
          <xsl:with-param name="str" select="$str_norm"/>
          <xsl:with-param name="prefix" select="'Q'"/>
        </xsl:call-template>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  <!--
  Function template for cleansing names of direct apply question fields.
  -->
  <xsl:template name="questionGroupName">
    <xsl:param name="str"/>
    <xsl:call-template name="elementName">
      <xsl:with-param name="str" select="$str"/>
      <xsl:with-param name="prefix" select="'Q'"/>
    </xsl:call-template>
  </xsl:template>
  <!--
  Function template for cleansing names of direct apply question fields.
  -->
  <xsl:template name="questionFieldName">
    <xsl:param name="str"/>
    <xsl:call-template name="elementName">
      <xsl:with-param name="str" select="$str"/>
      <xsl:with-param name="prefix" select="'Q'"/>
    </xsl:call-template>
  </xsl:template>
  <!--
  Function template for cleansing XML element names.
  -->
  <xsl:template name="elementName">
    <xsl:param name="str"/>
    <xsl:param name="prefix"/>
    <xsl:variable name="str_norm" select="translate($str, '&#x09;&#x0A;&#x0D;&#x20;', '')"/>
    <xsl:variable name="str_char" select="substring($str_norm, 1, 1)"/>
    <xsl:choose>
      <xsl:when test="number($str_char) = $str_char">
        <xsl:value-of select="concat($prefix, $str_norm)"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$str_norm"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
</xsl:stylesheet>
