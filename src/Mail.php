<?php

namespace Tualo\Office\MSGraphMail;

use Tualo\Office\Basic\TualoApplication as App;
use Tualo\Office\MSGraphVFS\StreamWrapper\SharePointStreamWrapper;
use Tualo\Office\MSGraph\API;

use Microsoft\Graph\Generated\Models;
use Microsoft\Graph\Generated\Models\User;
use Microsoft\Graph\Generated\Users\Item\MailFolders\Item\Messages\MessagesRequestBuilderGetQueryParameters;
use Microsoft\Graph\Generated\Users\Item\MailFolders\Item\Messages\MessagesRequestBuilderGetRequestConfiguration;
use Microsoft\Graph\Generated\Users\Item\SendMail\SendMailPostRequestBody;
use Microsoft\Graph\Generated\Users\Item\UserItemRequestBuilderGetQueryParameters;
use Microsoft\Graph\Generated\Users\Item\UserItemRequestBuilderGetRequestConfiguration;
use Microsoft\Graph\GraphRequestAdapter;
use Microsoft\Graph\GraphServiceClient;
use Microsoft\Kiota\Abstractions\Authentication\BaseBearerTokenAuthenticationProvider;
use Microsoft\Graph\Generated\Models\Message;
use Microsoft\Graph\Generated\Models\ItemBody;
use Microsoft\Graph\Generated\Models\BodyType;
use Microsoft\Graph\Generated\Models\Recipient;
use Microsoft\Graph\Generated\Models\EmailAddress;
use Microsoft\Graph\Generated\Models\FileAttachment;
use Microsoft\Kiota\Abstractions\ApiException;
use Microsoft\Graph\Generated\Models\ODataErrors\ODataError;

// require_once 'DeviceCodeTokenProvider.php';
use Microsoft\Kiota\Authentication\Oauth\ClientCredentialContext;
use Microsoft\Graph\Core\Authentication\GraphPhpLeagueAuthenticationProvider;
use Microsoft\Kiota\Abstractions\Authentication\AccessTokenProvider;

class Mail
{
    private static $config = [];

        public static function getInbox(): MessageCollectionResponse
    {
        $graphClient = API::GraphClient();

        if (API::has('clientSecret')) {
            return null;
        }
        
        $configuration = new MessagesRequestBuilderGetRequestConfiguration();
        $configuration->queryParameters = new MessagesRequestBuilderGetQueryParameters();
        // Only request specific properties
        $configuration->queryParameters->select = ['from', 'isRead', 'receivedDateTime', 'subject'];
        // Sort by received time, newest first
        $configuration->queryParameters->orderby = ['receivedDateTime DESC'];
        // Get at most 25 results
        $configuration->queryParameters->top = 25;
        return $graphClient->me()
            ->mailFolders()
            ->byMailFolderId('inbox')
            ->messages()
            ->get($configuration)->wait();
    }

 
    public static function sendMail(
        string $subject,
        string $bodyText,
        string $bodyHtml,
        string $recipient,
        array $attachments = [],
        string $listUnsubscribePost = ""
    ): void {

        try {
            $graphClient = API::GraphClient();

            if (!API::has('mailFromAddress')) {
                throw new \RuntimeException('Mail from address is not configured.');
            }

            
            $requestBody = new SendMailPostRequestBody();
            $message = new Message();
            $message->setSubject($subject);

            $message->setFrom(
                (new Recipient())
                    ->setEmailAddress(
                        (new EmailAddress())
                            ->setAddress(API::env('mailFromAddress'))
                    )
            );

            if ($bodyText != '') {
                $messageBody = new ItemBody();
                $messageBody->setContentType(new BodyType('text'));
                $messageBody->setContent($bodyText);
                $message->setBody($messageBody);
            }

            if ($bodyHtml != '') {
                $messageBody = new ItemBody();
                $messageBody->setContentType(new BodyType('html'));
                $messageBody->setContent($bodyHtml);
                $message->setBody($messageBody);
            }




            $toRecipientsRecipient1 = new Recipient();
            $toRecipientsRecipient1EmailAddress = new EmailAddress();
            $toRecipientsRecipient1EmailAddress->setAddress($recipient);
            $toRecipientsRecipient1->setEmailAddress($toRecipientsRecipient1EmailAddress);
            $toRecipientsArray[] = $toRecipientsRecipient1;


            $attachmentsArray = [];
            foreach ($attachments as $attachment) {


                $attachmentsAttachment1 = new FileAttachment();
                if (isset($attachment['isInline'])) {
                    $attachmentsAttachment1->setIsInline($attachment['isInline']);
                }
                $attachmentsAttachment1->setName($attachment['name']);
                if (isset($attachment['isInline'])) {
                    $attachmentsAttachment1->setContentType($attachment['contentType']);
                }
                if (isset($attachment['content'])) {
                    $attachmentsAttachment1->setContentBytes(\GuzzleHttp\Psr7\Utils::streamFor(base64_encode($attachment['content'])));
                }
                $attachmentsArray[] = $attachmentsAttachment1;
            }
            if (count($attachmentsArray) > 0)
                $message->setAttachments($attachmentsArray);

            $message->setToRecipients($toRecipientsArray);


            // Füge den List-Unsubscribe-Post-Header hinzu, falls angegeben
            if ($listUnsubscribePost !== null) {
                $extendedProperty = new Models\SingleValueLegacyExtendedProperty();
                $extendedProperty->setId("String 0x1045"); // Standard-ID für benutzerdefinierte Header
                $extendedProperty->setValue($listUnsubscribePost);
                $message->setSingleValueExtendedProperties([$extendedProperty]);
            }

            $requestBody->setMessage($message);
            $graphClient->me()->sendMail()->post($requestBody)->wait();
        } catch (ODataError $e) {
            echo $e->getError()->getMessage();
            throw new \Exception($e->getError()->getMessage());
        } catch (ApiException $ex) {
            echo $ex->getMessage();
        } catch (Exception $ex) {
            echo $ex->getMessage();
        }
    }

}
