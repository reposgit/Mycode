<?php
add_action( 'wp_enqueue_scripts', 'theme_enqueue_styles' );
function theme_enqueue_styles() {
    wp_enqueue_style( 'parent-style', get_template_directory_uri() . '/style.css' );
    wp_enqueue_style( 'et-font-awesome',get_stylesheet_directory_uri().'/css/font-awesome.css', array( 'fonts' ) );
}

// Hook in
add_filter( 'woocommerce_checkout_fields' , 'custom_override_checkout_fields' );

// Our hooked in function - $fields is passed via the filter!
function custom_override_checkout_fields( $fields ) {
     unset($fields['order']['order_comments']);

     return $fields;
}

add_action( 'wp_enqueue_scripts', 'my_scripts_method' );
function my_scripts_method(){
  wp_enqueue_script( 'newscript', get_template_directory_uri() . '/js/jquery.maskedinput.js');
}



add_action( 'woocommerce_checkout_order_processed', 'pending_new_order_notification', 20, 1 );
function pending_new_order_notification( $order_id ) {

    $order = wc_get_order( $order_id );

    if( ! $order->has_status( 'pending' ) ) return;

    $wc_email = WC()->mailer()->get_emails()['WC_Email_New_Order'];

    ## -- Настройка заголовка, темы (и при желании добавить получателей)  -- ##
    // Изменяем тему
    $wc_email->settings['subject'] = __('{site_title} - New customer Pending order ({order_number}) - {order_date}');
    // Изменяем заголовок
    $wc_email->settings['heading'] = __('New customer Pending Order');
    
    // Отправить уведомление «Новое письмо» (администратору)
    $wc_email->trigger( $order_id );
    // Сообщение пользователю
         $email_heading = 'Спасибо за заказ';
            
            $args = array(
                        'order'         => $order,
                        'email_heading' => $email_heading,
                        'sent_to_admin' => false,
                        'plain_text'    => false,
                    );
            $content_info = wc_get_template_html("emails/customer-processing-order-my.php", $args);
            
            $site_title = 'Cacao Club';
            $customer_email = $order->get_billing_email();
            $email_subject = 'Cacao Club - заказ №'.$order->get_order_number();
            wc_mail($customer_email, $email_subject, $content_info);
    }

add_filter( 'woocommerce_cart_subtotal', 'slash_cart_subtotal_if_discount', 99, 3 );
function slash_cart_subtotal_if_discount( $cart_subtotal, $compound, $obj ){
global $woocommerce;
if ( $woocommerce->cart->get_cart_discount_total() <> 0 ) {
$new_cart_subtotal = wc_price( WC()->cart->subtotal - $woocommerce->cart->get_cart_discount_tax_total() - $woocommerce->cart->get_cart_discount_total() );
$cart_subtotal = sprintf( '<del>%s</del> <b>%s</b>', $cart_subtotal, $new_cart_subtotal );
}
return $cart_subtotal;
}

add_action('woocommerce_checkout_process', 'wh_phoneValidateCheckoutFields');


function wh_phoneValidateCheckoutFields() {
$billing_phone = filter_input(INPUT_POST, 'billing_phone');
if ($billing_phone[4] != '9')
{
wc_add_notice(__('<strong>Мобильные номера РФ начинаются с 9, если вы из другой страны свяжитесь с нами</strong>'), 'error');    
}
if (strlen($billing_phone) < 18) {
wc_add_notice(__('<strong>Проверьте мобильный номер</strong>'), 'error');
}
}