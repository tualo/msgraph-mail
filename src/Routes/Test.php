<?php

namespace Tualo\Office\MSGraphMail\Routes;

use Tualo\Office\Basic\TualoApplication as App;
use Tualo\Office\Basic\Route as BasicRoute;
use Tualo\Office\MSGraphMail\Mail ;
use Tualo\Office\MSGraph\API;

class Test extends \Tualo\Office\Basic\RouteWrapper
{
    public static function register()
    {
        BasicRoute::add('/msgraph-mail/test', function ($matches) {
            App::contenttype('application/json');

            try {
                $to = App::configuration('msgraph-mail', 'testmail', '');
                if (empty($to)) {
                    throw new \RuntimeException('Testmail address is not configured.');
                }
                Mail::sendMail(
                    'Testmail',
                    'Dies ist ein Testmail',
                    'Dies ist ein Testmail',
                    $to,
                    [],
                    ''
                );
                App::result('success', true);
            } catch (\Exception $e) {
                App::result('error', $e->getMessage());
            }
        }, ['get'], true);
    }
}
